#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH
HỆ THỐNG TỔNG HỢP VÀ XỬ LÝ DỮ LIỆU NTQLAHS (CHẠY OFFLINE 100%)
------------------------------------------------------------------------------
Châm ngôn: "CÔNG MINH - CHÍNH TRỰC - KHÁCH QUAN - THẬN TRỌNG - KHIÊM TỐN"

Chức năng:
- Đọc file Excel (.xlsx) copy từ phần mềm Quản lý án hình sự (QLAHS).
- Hoặc nhận dữ liệu dán trực tiếp (Ctrl+V) từ clipboard.
- Tự động bóc tách và chuẩn hóa thành bảng 12 cột chuẩn:
  1. STT
  2. Mã vụ án
  3. Tên vụ án
  4. Cơ quan cảnh sát điều tra
  5. Tỉnh/Thành (Tự động đọc và nhận dạng từ CQĐT)
  6. Giai đoạn
  7. Mã bị can
  8. Tên bị can
  9. Ngày sinh
  10. Số CCCD
  11. Điều luật
  12. Trạng thái hồ sơ (Đình nã, Truy nã, Tạm đình chỉ...)
- Sinh bảng Pivot Table ma trận tổng quan: Tỉnh/Thành × Giai đoạn / Trạng thái hồ sơ.
- Giữ nguyên số 0 ở đầu cho Mã vụ án, Mã bị can, Số CCCD (Text '@').
- Xuất file Excel chuẩn gồm 2 Sheet:
  * Sheet 1: Danh_Sach_Chuan_Hoa (12 cột)
  * Sheet 2: Tong_Hop_Pivot_Table (Ma trận tổng hợp theo Tỉnh/Thành và Giai đoạn)
"""

import json
import os
import queue
import re
import sys
import threading
import time
from datetime import datetime, date
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BarChart, PieChart, Reference
except ImportError:
    openpyxl = None
    BarChart = None
    PieChart = None
    Reference = None


def get_resource_path(relative_path):
    """Lấy đường dẫn tài nguyên tĩnh tương thích cả khi chạy script và khi đóng gói PyInstaller OneFile."""
    try:
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    p = os.path.join(base_path, relative_path)
    if os.path.exists(p):
        return p
    try:
        exe_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
        p2 = os.path.join(exe_dir, relative_path)
        if os.path.exists(p2):
            return p2
    except Exception:
        pass
    return p


# ==============================================================================
# HỆ THỐNG CẤU HÌNH TÙY BIẾN ĐA NĂNG (SETTINGS & PREFERENCES)
# ==============================================================================
DEFAULT_APP_CONFIG = {
    "agency_name": "VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH",
    "author_title": "Trần Văn Khoa - Kiểm tra viên / VKSND TP.HCM",
    "warning_thresholds": {
        "urgent_days": 30,
        "very_urgent_days": 60,
        "warning_days": 90
    },
    "unit_aliases": {
        "P2": "Phòng 2 - VKSND TP.HCM",
        "P.2": "Phòng 2 - VKSND TP.HCM",
        "Phòng 2": "Phòng 2 - VKSND TP.HCM",
        "P1": "Phòng 1 - VKSND TP.HCM",
        "P.1": "Phòng 1 - VKSND TP.HCM",
        "Phòng 1": "Phòng 1 - VKSND TP.HCM",
        "P3": "Phòng 3 - VKSND TP.HCM",
        "P.3": "Phòng 3 - VKSND TP.HCM",
        "Phòng 3": "Phòng 3 - VKSND TP.HCM"
    }
}

def get_config_filepath():
    """Lấy đường dẫn tệp cấu hình JSON luôn nằm cạnh tệp .exe hoặc script."""
    try:
        if getattr(sys, 'frozen', False):
            base = os.path.dirname(sys.executable)
        else:
            base = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base, "config_qlahs.json")
    except Exception:
        return "config_qlahs.json"

def load_app_config():
    """Tải cấu hình người dùng từ tệp config_qlahs.json, tự động hợp nhất với cấu hình mặc định."""
    p = get_config_filepath()
    if os.path.exists(p):
        try:
            with open(p, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                res = dict(DEFAULT_APP_CONFIG)
                res.update(cfg)
                if "warning_thresholds" in cfg:
                    res["warning_thresholds"] = dict(DEFAULT_APP_CONFIG["warning_thresholds"])
                    res["warning_thresholds"].update(cfg["warning_thresholds"])
                if "unit_aliases" in cfg:
                    res["unit_aliases"] = dict(DEFAULT_APP_CONFIG["unit_aliases"])
                    res["unit_aliases"].update(cfg["unit_aliases"])
                return res
        except Exception:
            pass
    return dict(DEFAULT_APP_CONFIG)

def save_app_config(cfg):
    """Lưu cấu hình người dùng vào tệp config_qlahs.json."""
    p = get_config_filepath()
    try:
        with open(p, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False

CURRENT_CONFIG = load_app_config()

# Danh mục nhận dạng 63 Tỉnh/Thành phố & Cơ quan Trung ương
PROVINCE_RULES = [
    # 5 Thành phố trực thuộc Trung ương
    ("TP. Hồ Chí Minh", ["hồ chí minh", "ho chi minh", "tphcm", "tp.hcm", "tp hcm", "sài gòn", "thủ đức", "quận 1", "quận 3", "quận 4", "quận 5", "quận 6", "quận 7", "quận 8", "quận 10", "quận 11", "quận 12", "bình thạnh", "gò vấp", "tân bình", "tân phú", "phú nhuận", "bình tân", "củ chi", "hóc môn", "bình chánh", "nhà bè", "cần giờ"]),
    ("Hà Nội", ["hà nội", "ha noi", "thủ đô hà nội", "hn", "ba đình", "hoàn kiếm", "đống đa", "hai bà trưng", "cầu giấy", "thanh xuân", "hoàng mai", "long biên", "nam từ liêm", "bắc từ liêm", "tây hồ", "hà đông", "sơn tây", "ba vì", "chương mỹ", "đan phượng", "đông anh", "gia lâm", "hoài đức", "mê linh", "mỹ đức", "phú xuyên", "phúc thọ", "quốc oai", "sóc sơn", "thạch thất", "thanh oai", "thanh trì", "thường tín", "ứng hòa"]),
    ("Hải Phòng", ["hải phòng", "hai phong", "hp", "ngô quyền", "hồng bàng", "lê chân", "hải an", "kiến an", "đồ sơn", "dương kinh", "thủy nguyên", "an dương", "an lão", "kiến thụy", "tiên lãng", "vĩnh bảo", "cát hải", "bạch long vĩ"]),
    ("Đà Nẵng", ["đà nẵng", "da nang", "hải châu", "thanh khê", "sơn trà", "ngũ hành sơn", "liên chiểu", "cẩm lệ", "hòa vang", "hoàng sa"]),
    ("Cần Thơ", ["cần thơ", "can tho", "ninh kiều", "bình thủy", "cái răng", "ô môn", "thốt nốt", "phong điền", "cờ đỏ", "thới lai", "vĩnh thạnh"]),

    # Miền Bắc
    ("Bắc Giang", ["bắc giang", "bac giang", "việt yên", "hiệp hòa", "lạng giang", "lục nam", "lục ngạn", "sơn động", "tân yên", "yên dũng", "yên thế"]),
    ("Bắc Kạn", ["bắc kạn", "bắc cạn", "bac kan", "bac can", "ba bể", "bạch thông", "chợ đồn", "chợ mới", "na rì", "ngân sơn", "pác nặm"]),
    ("Bắc Ninh", ["bắc ninh", "bac ninh", "từ sơn", "thuận thành", "quế võ", "yên phong", "tiên du", "gia bình", "lương tài"]),
    ("Cao Bằng", ["cao bằng", "cao bang", "bảo lạc", "bảo lâm", "hạ lang", "hà quảng", "hòa an", "nguyên bình", "thạch an", "trùng khánh"]),
    ("Điện Biên", ["điện biên", "dien bien", "mường lay", "mường nhé", "mường chà", "tủa chùa", "tuần giáo", "điện biên đông", "nậm pồ"]),
    ("Hà Giang", ["hà giang", "ha giang", "đồng văn", "mèo vạc", "yên minh", "quản bạ", "vị xuyên", "bắc mê", "hoàng su phì", "xín mần", "bắc quang", "quang bình"]),
    ("Hà Nam", ["hà nam", "ha nam", "phủ lý", "duy tiên", "kim bảng", "thanh liêm", "bình lục", "lý nhân"]),
    ("Hải Dương", ["hải dương", "hai duong", "chí linh", "kinh môn", "bình giang", "cẩm giàng", "gia lộc", "kim thành", "nam sách", "ninh giang", "thanh hà", "thanh miện", "tứ kỳ"]),
    ("Hòa Bình", ["hòa bình", "hoa binh", "lương sơn", "cao phong", "đà bắc", "kim bôi", "lạc sơn", "lạc thủy", "mai châu", "tân lạc", "yên thủy"]),
    ("Hưng Yên", ["hưng yên", "hung yen", "mỹ hào", "ân thi", "khoái châu", "kim động", "phù cừ", "tiên lữ", "văn giang", "văn lâm", "yên mỹ"]),
    ("Lai Châu", ["lai châu", "lai chau", "mường tè", "sìn hồ", "nậm nhùn", "tam đường", "tân uyên", "than uyên", "phong thổ"]),
    ("Lạng Sơn", ["lạng sơn", "lang son", "bắc sơn", "bình gia", "cao lộc", "chi lăng", "đình lập", "hữu lũng", "lộc bình", "tràng định", "văn lãng", "văn quan"]),
    ("Lào Cai", ["lào cai", "lao cai", "sa pa", "bát xát", "bảo thắng", "bảo yên", "bắc hà", "mường khương", "si ma cai", "văn bàn"]),
    ("Nam Định", ["nam định", "nam dinh", "giao thủy", "hải hậu", "mỹ lộc", "nam trực", "nghĩa hưng", "trực ninh", "vụ bản", "xuân trường", "ý yên"]),
    ("Ninh Bình", ["ninh bình", "ninh binh", "tam điệp", "gia viễn", "hoa lư", "kim sơn", "nho quan", "yên khánh", "yên mô"]),
    ("Phú Thọ", ["phú thọ", "phu tho", "việt trì", "cẩm khê", "đoan hùng", "hạ hòa", "lâm thao", "phù ninh", "tam nông", "tân sơn", "thanh ba", "thanh sơn", "thanh thủy", "yên lập"]),
    ("Quảng Ninh", ["quảng ninh", "quang ninh", "hạ long", "cẩm phả", "uông bí", "móng cái", "đông triều", "quảng yên", "vân đồn", "tiên yên", "ba chẽ", "bình liêu", "đầm hà", "hải hà", "cô tô"]),
    ("Sơn La", ["sơn la", "son la", "mộc châu", "quỳnh nhai", "mường la", "thuận châu", "phù yên", "bắc yên", "mai sơn", "sông mã", "sốp cộp", "vân hồ"]),
    ("Thái Bình", ["thái bình", "thai binh", "đông hưng", "hưng hà", "kiến xương", "quỳnh phụ", "thái thụy", "tiền hải", "vũ thư"]),
    ("Thái Nguyên", ["thái nguyên", "thai nguyen", "sông công", "phổ yên", "đại từ", "định hóa", "đồng hỷ", "phú bình", "phú lương", "võ nhai"]),
    ("Tuyên Quang", ["tuyên quang", "tuyen quang", "chiêm hóa", "hàm yên", "lâm bình", "na hang", "sơn dương", "yên sơn"]),
    ("Vĩnh Phúc", ["vĩnh phúc", "vinh phuc", "vĩnh yên", "phúc yên", "bình xuyên", "lập thạch", "sông lô", "tam đảo", "tam dương", "vĩnh tường", "yên lạc"]),
    ("Yên Bái", ["yên bái", "yen bai", "nghĩa lộ", "lục yên", "mù cang chải", "trạm tấu", "trấn yên", "văn chấn", "văn yên", "yên bình"]),

    # Miền Trung & Tây Nguyên
    ("Hà Tĩnh", ["hà tĩnh", "ha tinh", "hồng lĩnh", "kỳ anh", "hương sơn", "cẩm xuyên", "nghi xuân", "đức thọ", "hương khê", "lộc hà", "thạch hà", "can lộc", "vũ quang"]),
    ("Nghệ An", ["nghệ an", "nghe an", "vinh", "thái hòa", "cửa lò", "hoàng mai", "diễn châu", "quỳnh lưu", "yên thành", "đô lương", "nghi lộc", "nam đàn", "hưng nguyên", "thanh chương", "anh sơn", "con cuông", "tương dương", "kỳ sơn", "quỳ hợp", "quỳ châu", "quế phong", "tân kỳ"]),
    ("Thanh Hóa", ["thanh hóa", "thanh hoa", "sầm sơn", "bỉm sơn", "nghi sơn", "bá thước", "cẩm thủy", "đông sơn", "hà trung", "hậu lộc", "hoằng hóa", "lang chánh", "mường lát", "nga sơn", "ngọc lặc", "như thanh", "như xuân", "nông cống", "quan hóa", "quan sơn", "quảng xương", "thạch thành", "thiệu hóa", "thọ xuân", "thường xuân", "triệu sơn", "vĩnh lộc", "yên định"]),
    ("Quảng Bình", ["quảng bình", "quang binh", "đồng hới", "ba đồn", "bố trạch", "lệ thủy", "minh hóa", "quảng ninh", "quảng trạch", "tuyên hóa"]),
    ("Quảng Trị", ["quảng trị", "quang tri", "đông hà", "quảng trị", "cam lộ", "cồn cỏ", "đa krông", "gio linh", "hướng hóa", "hải lăng", "triệu phong", "vĩnh linh"]),
    ("Thừa Thiên Huế", ["thừa thiên huế", "thừa thiên - huế", "thừa thiên", "thua thien hue", "huế", "hương thủy", "hương trà", "a lưới", "nam đông", "phong điền", "phú lộc", "phú vang", "quảng điền"]),
    ("Quảng Nam", ["quảng nam", "quang nam", "tam kỳ", "hội an", "điện bàn", "bắc trà my", "nam trà my", "đại lộc", "đông giang", "duy xuyên", "hiệp đức", "nam giang", "nông sơn", "núi thành", "phú ninh", "phước sơn", "quế sơn", "tây giang", "thăng bình", "tiên phước"]),
    ("Quảng Ngãi", ["quảng ngãi", "quang ngai", "đức phổ", "ba tơ", "bình sơn", "lý sơn", "minh long", "mộ đức", "nghĩa hành", "sơn hà", "sơn tây", "sơn tịnh", "trà bồng", "tư nghĩa"]),
    ("Bình Định", ["bình định", "binh dinh", "quy nhơn", "an nhơn", "hoài nhơn", "an lão", "hoài ân", "phù cát", "phù mỹ", "tây sơn", "tuy phước", "vân canh", "vĩnh thạnh"]),
    ("Phú Yên", ["phú yên", "phu yen", "tuy hòa", "sông cầu", "đông hòa", "đồng xuân", "phú hòa", "sơn hòa", "sông hinh", "tây hòa", "tuy an"]),
    ("Khánh Hòa", ["khánh hòa", "khanh hoa", "nha trang", "cam ranh", "ninh hòa", "cam lâm", "diên khánh", "khánh sơn", "khánh vĩnh", "trường sa", "vạn ninh"]),
    ("Ninh Thuận", ["ninh thuận", "ninh thuan", "phan rang", "tháp chàm", "bác ái", "ninh hải", "ninh phước", "ninh sơn", "thuận bắc", "thuận nam"]),
    ("Bình Thuận", ["bình thuận", "binh thuan", "phan thiết", "la gi", "bắc bình", "đức linh", "hàm tân", "hàm thuận bắc", "hàm thuận nam", "phú quý", "tánh linh", "tuy phong"]),
    ("Kon Tum", ["kon tum", "kontum", "đăk hà", "đăk tô", "ia h'drai", "kon plông", "kon rẫy", "ngọc hồi", "sa thầy", "tu mơ rông"]),
    ("Gia Lai", ["gia lai", "pleiku", "an khê", "ayar pa", "chư păh", "chư prông", "chư pưh", "chư sê", "đăk đoa", "đăk pơ", "đức cơ", "ia grai", "ia pa", "k'bang", "kông chro", "krông pa", "mang yang", "phú thiện"]),
    ("Đắk Lắk", ["đắk lắk", "đắc lắc", "đăk lăk", "dak lak", "daklak", "buôn ma thuột", "buôn hồ", "buôn đôn", "cư kuin", "cư m'gar", "ea h'leo", "ea kar", "ea súp", "krông ana", "krông bông", "krông bắk", "krông năng", "krông pắc", "lắk", "m'đrắk"]),
    ("Đắk Nông", ["đắk nông", "đắc nông", "đăk nông", "dak nong", "gia nghĩa", "cư jút", "đắk glong", "đắk mil", "đắk r'lấp", "đắk song", "krông nô", "tuy đức"]),
    ("Lâm Đồng", ["lâm đồng", "lam dong", "đà lạt", "bảo lộc", "bảo lâm", "cát tiên", "di linh", "đạ hoai", "đạ tẻh", "đam rông", "đơn dương", "đức trọng", "lạc dương", "lâm hà"]),

    # Miền Đông & Tây Nam Bộ
    ("Bình Dương", ["bình dương", "binh duong", "thủ dầu một", "dĩ an", "thuận an", "bến cát", "tân uyên", "bàu bàng", "dầu tiếng", "phú giáo", "bắc tân uyên"]),
    ("Bình Phước", ["bình phước", "binh phuoc", "đồng xoài", "bình long", "phước long", "bù đăng", "bù đốp", "bù gia mập", "chơn thành", "đồng phú", "hớn quản", "lộc ninh", "phú riêng"]),
    ("Đồng Nai", ["đồng nai", "dong nai", "biên hòa", "long khánh", "cẩm mỹ", "định quán", "long thành", "nhơn trạch", "tân phú", "thống nhất", "trảng bom", "vĩnh cửu", "xuân lộc"]),
    ("Tây Ninh", ["tây ninh", "tay ninh", "trảng bàng", "hòa thành", "bến cầu", "châu thành", "dương minh châu", "gò dầu", "tân biên", "tân châu"]),
    ("Bà Rịa - Vũng Tàu", ["bà rịa", "vũng tàu", "ba ria", "vung tau", "brvt", "br-vt", "phú mỹ", "châu đức", "côn đảo", "đất đỏ", "long điền", "xuyên mộc"]),
    ("Long An", ["long an", "tân an", "kiến tường", "bến lức", "cần đước", "cần giuộc", "châu thành", "đức hòa", "đức huệ", "mộc hóa", "tân hưng", "tân thạnh", "tân trụ", "thạnh hóa", "thủ thừa", "vĩnh hưng"]),
    ("Tiền Giang", ["tiền giang", "tien giang", "mỹ tho", "gò công", "cai lậy", "cái bè", "châu thành", "chợ gạo", "tân phước", "tân phú đông"]),
    ("Bến Tre", ["bến tre", "ben tre", "ba tri", "bình đại", "châu thành", "chợ lách", "giồng trôm", "mỏ cày bắc", "mỏ cày nam", "thạnh phú"]),
    ("Trà Vinh", ["trà vinh", "tra vinh", "duyên hải", "càng long", "cầu kè", "cầu ngang", "châu thành", "tiểu cần", "trà cú"]),
    ("Vĩnh Long", ["vĩnh long", "vinh long", "bình minh", "bình tân", "long hồ", "mang thít", "tam bình", "trà ôn", "vũng liêm"]),
    ("Đồng Tháp", ["đồng tháp", "dong thap", "cao lãnh", "sa đéc", "hồng ngự", "châu thành", "lai vung", "lấp vò", "tam nông", "tân hồng", "thanh bình", "tháp mười"]),
    ("An Giang", ["an giang", "long xuyên", "châu đốc", "tân châu", "tịnh biên", "an phú", "châu phú", "châu thành", "chợ mới", "phú tân", "thoại sơn", "tri tôn"]),
    ("Kiên Giang", ["kiên giang", "kien giang", "rạch giá", "hà tiên", "phú quốc", "an biên", "an minh", "châu thành", "giang thành", "giồng riêng", "gò quao", "hòn đất", "kiên hải", "kiên lương", "tân hiệp", "u minh thượng", "vĩnh thuận"]),
    ("Hậu Giang", ["hậu giang", "hau giang", "vị thanh", "ngã bảy", "châu thành", "châu thành a", "long mỹ", "phụng hiệp", "vị thủy"]),
    ("Sóc Trăng", ["sóc trăng", "soc trang", "ngã năm", "vĩnh châu", "châu thành", "cù lao dung", "kế sách", "mỹ tú", "mỹ xuyên", "thạnh trị", "trần đề"]),
    ("Bạc Liêu", ["bạc liêu", "bac lieu", "giá rai", "đông hải", "hòa bình", "hồng dân", "phước long", "vĩnh lợi"]),
    ("Cà Mau", ["cà mau", "ca mau", "cái nước", "đầm dơi", "năm căn", "ngọc hiển", "phú tân", "thới bình", "trần văn thời", "u minh"]),

    # Cơ quan Trung ương / Bộ ngành
    ("Bộ Công an", ["bộ công an", "cục c01", "cục c02", "cục c03", "cục c04", "cục c05", "cục a09", "cục cảnh sát điều tra", "cơ quan an ninh điều tra bộ công an", "cục an ninh", "bca"]),
    ("VKSND tối cao", ["vksnd tối cao", "vksndtc", "viện kiểm sát nhân dân tối cao", "cơ quan điều tra vksnd", "cục 1 vksnd"]),
    ("Bộ Quốc phòng", ["bộ quốc phòng", "cục điều tra hình sự bộ quốc phòng", "quân khu", "bộ tư lệnh"]),
]

from datetime import datetime, date

def parse_date_safely(val):
    """Chuyển đổi chuỗi ngày tháng hoặc đối tượng datetime sang date chuẩn an toàn."""
    if not val:
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    val_s = str(val).replace("T", " ").strip()
    if not val_s:
        return None

    for fmt in (
        '%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M',
        '%d/%m/%Y %H:%M:%S', '%d/%m/%Y %H:%M',
        '%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%d.%m.%Y',
        '%d/%m/%y', '%Y/%m/%d'
    ):
        try:
            return datetime.strptime(val_s, fmt).date()
        except ValueError:
            pass

    m_dmy = re.search(r'\b(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})\b', val_s)
    if m_dmy:
        d, m, y = int(m_dmy.group(1)), int(m_dmy.group(2)), int(m_dmy.group(3))
        try:
            return date(y, m, d)
        except ValueError:
            pass

    m_ymd = re.search(r'\b(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})\b', val_s)
    if m_ymd:
        y, m, d = int(m_ymd.group(1)), int(m_ymd.group(2)), int(m_ymd.group(3))
        try:
            return date(y, m, d)
        except ValueError:
            pass

    return None


def format_date_str(val):
    """Định dạng chuỗi ngày tháng sang chuẩn Việt Nam DD/MM/YYYY."""
    if not val:
        return ""
    d = parse_date_safely(val)
    if d:
        return d.strftime("%d/%m/%Y")
    return str(val).strip()

CQDT_PREFIXES = (
    "công an",
    "cơ quan",
    "cqcsđt",
    "cqđt",
    "cục điều tra",
    "viện kiểm sát",
    "bộ tư lệnh",
    "chi cục",
    "đồn biên phòng",
    "hải quan",
    "kiểm lâm",
)

COLUMNS_12 = [
    ("stt", "STT", 50),
    ("ma_vu_an", "Mã vụ án", 160),
    ("ten_vu_an", "Tên vụ án", 220),
    ("cqdt", "Cơ quan cảnh sát điều tra", 180),
    ("tinh_thanh", "Tỉnh/Thành", 130),
    ("giai_doan", "Giai đoạn", 95),
    ("ma_bi_can", "Mã bị can", 165),
    ("ten_bi_can", "Tên bị can", 145),
    ("ngay_sinh", "Ngày sinh", 95),
    ("so_cccd", "Số CCCD", 115),
    ("dieu_luat", "Điều luật", 220),
    ("trang_thai", "Trạng thái hồ sơ", 120),
]

# 26 Cột chuẩn hóa cho định dạng Sổ Vụ án / Sổ Thụ lý / Thời hiệu truy cứu TNHS
# Được phân nhóm khoa học: Vụ án -> Khởi tố -> Thụ lý VKS -> Tiến độ & Thời hiệu -> Bị can & Tội danh -> Phân công & Kết quả

# Danh sách lãnh đạo Viện và lãnh đạo Phòng theo phân công của VKSND TP.HCM
LANH_DAO_VIEN = [
    "Bùi Xuân Thái",
    "Võ Quang Huy",
]

LANH_DAO_PHONG_2 = [
    "Nguyễn Quỳnh Lan",
    "Đỗ Tất Minh Tuấn",
    "Đào Thị Hồng Vân",
    "Nguyễn Duy Thạch",
    "Trần Tấn Lập",
    "Nguyễn Quang Duyệt",
    "Cao Xuân Cường",
]


def is_procedural_status_text(text):
    """
    Kiểm tra xem một đoạn văn bản có phải là trạng thái tố tụng / trạng thái vụ án thay vì tên cán bộ.
    Giúp ngăn chặn tuyệt đối tình trạng hiểu lầm trạng thái vụ án thành Kiểm sát viên giải quyết.
    """
    if not text:
        return False
    t = str(text).lower().strip()
    
    # 1. Các tiền tố giai đoạn tố tụng
    if re.search(r'^(?:điều tra|truy tố|sơ thẩm|phúc thẩm|thi hành án|thụ lý|tiếp nhận)[\s.:\-_/]+', t):
        return True
        
    # 2. Các cụm từ nghiệp vụ về tiến độ, trạng thái giải quyết đặc trưng
    status_keywords = (
        "thqct", "thực hành quyền công tố", "kiểm sát điều tra", "kiểm sát xét xử",
        "chuyển để nhập", "nhập vụ án", "tách vụ án", "chuyển vụ án", "chuyển cơ quan khác",
        "chuyển giải quyết", "chờ giải quyết", "chờ nhập vahs", "chờ cấp trên", "chờ chuyển",
        "tạm đình chỉ", "đình chỉ vụ án", "đình chỉ điều tra", "phục hồi giải quyết",
        "lưu hành nội bộ", "phát hành văn bản", "chờ phát hành", "chưa phát hành", "đã phát hành",
        "biểu mẫu", "hồ sơ vụ án", "thời hạn giải quyết", "hết thời hiệu", "thụ lý mới",
        "đang giải quyết", "sau xét xử", "đã xét xử"
    )
    if any(kw in t for kw in status_keywords):
        return True
        
    # 3. Dạng gán số văn bản biểu mẫu (như ': 2', ': 10')
    if re.search(r':\s*\d+', t):
        return True

    return False


def is_form_status_text(text):
    """Kiểm tra văn bản trạng thái biểu mẫu (như 'Lưu hành nội bộ: 2', 'Đã phát hành văn bản: 1')."""
    if not text:
        return False
    t = str(text).lower().strip()
    return any(x in t for x in ("lưu hành nội bộ", "phát hành văn bản", "chờ phát hành", "chưa phát hành", "đã phát hành", "biểu mẫu"))


def clean_person_name(name):
    """
    Chuẩn hóa tên cán bộ/lãnh đạo:
    - Loại bỏ hoàn toàn các tiền tố chức danh: PVT, Phó Viện trưởng, LĐ Viện, TP, PTP, KSV, CV, Đ/c, Đc...
    - Loại bỏ các ký tự thừa (chấm, hai chấm, gạch ngang...).
    - Xử lý lặp từ ở cuối (ví dụ 'Bùi Xuân Thái Thái' -> 'Bùi Xuân Thái').
    - Chuẩn hóa về đúng họ tên đích danh chính thức nếu thuộc danh sách Lãnh đạo Viện / Lãnh đạo Phòng.
    - Tuyệt đối loại bỏ các chuỗi trạng thái tố tụng / biểu mẫu nếu bị lọt vào.
    """
    if not name:
        return ""
    if is_procedural_status_text(name) or is_form_status_text(name):
        return ""
    s = str(name).replace("\xa0", " ").strip()
    
    # Bóc tách triệt để các tiền tố chức danh (lặp lại xử lý tiền tố kép như 'Đ/c PVT.')
    pattern_prefix = r'^(?:đ\/c|đc|lđ\s*viện|lãnh\s*đạo\s*viện|pvt|phó\s*viện\s*trưởng|viện\s*trưởng|trưởng\s*phòng|phó\s*trưởng\s*phòng|lđ\s*phòng|tp|ptp|ksv|kiểm\s*sát\s*viên|cv|chuyên\s*viên)[\s.:\-_/]*'
    while re.search(pattern_prefix, s, re.I):
        s = re.sub(pattern_prefix, '', s, flags=re.I).strip()
    
    s = re.sub(r'^[.\-:\s,;]+', '', s)
    s = re.sub(r'[.\-:\s,;]+$', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    
    if not s or is_procedural_status_text(s) or is_form_status_text(s):
        return ""

    # Xử lý lặp từ cuối (ví dụ 'Bùi Xuân Thái Thái' -> 'Bùi Xuân Thái')
    words = s.split()
    if len(words) >= 3 and words[-1].lower() == words[-2].lower():
        words.pop()
        s = " ".join(words)

    # Khớp đích danh Lãnh đạo Viện (trả về tên chuẩn không kèm chức vụ)
    s_low = s.lower()
    for ldv in LANH_DAO_VIEN:
        if ldv.lower() in s_low or s_low in ldv.lower():
            return ldv

    # Khớp đích danh Lãnh đạo Phòng 2 (trả về tên chuẩn không kèm chức vụ)
    for ldp in LANH_DAO_PHONG_2:
        if ldp.lower() in s_low or s_low in ldp.lower():
            return ldp

    return s


def split_officers(raw_text, can_bo_thu_ly=""):
    """
    Phan tách chính xác cán bộ thành 4 cột nghiệp vụ độc lập:
    1. lanh_dao_vien: Lãnh đạo Viện phụ trách (chỉ họ và tên, đã xóa mọi tiền tố PVT...)
    2. lanh_dao_phong: Lãnh đạo Phòng phụ trách (chỉ họ và tên, đã xóa mọi tiền tố TP/PTP...)
    3. can_bo_giai_quyet: Kiểm sát viên giải quyết (đã loại bỏ tên LĐ Viện và LĐ Phòng)
    4. can_bo_thu_ly: Cán bộ thụ lý / Chuyên viên giúp việc
    """
    if not raw_text and not can_bo_thu_ly:
        return "", "", "", ""

    text = str(raw_text or "").replace("\xa0", " ").strip()
    clean_text = re.sub(r'[\(\)]', ' ', text)
    raw_parts = [p.strip() for p in re.split(r'[,;|\n]+', clean_text) if p.strip()]

    ld_vien_found = []
    ld_phong_found = []
    ksv_found = []
    cv_found = []

    for item in raw_parts:
        if not item or len(item) < 2:
            continue
        if is_procedural_status_text(item) or is_form_status_text(item):
            continue
        
        item_lower = item.lower()

        # 1. Nhận diện chức danh viết tắt rõ ràng cho Lãnh đạo Viện
        if re.search(r'^(?:pvt|phó\s*viện\s*trưởng|lđ\s*viện)[\s.:\-_/]*', item_lower):
            c_name = clean_person_name(item)
            if c_name and c_name not in ld_vien_found:
                ld_vien_found.append(c_name)
            continue

        # 2. Nhận diện chức danh viết tắt cho Lãnh đạo Phòng
        if re.search(r'^(?:tp|ptp|trưởng\s*phòng|phó\s*trưởng\s*phòng)[\s.:\-_/]*', item_lower):
            c_name = clean_person_name(item)
            if c_name and c_name not in ld_phong_found:
                ld_phong_found.append(c_name)
            continue

        # 3. Nhận diện Chuyên viên / Cán bộ thụ lý
        if re.search(r'^(?:cv|chuyên\s*viên)[\s.:\-_/]*', item_lower):
            c_name = clean_person_name(item)
            if c_name and c_name not in cv_found:
                cv_found.append(c_name)
            continue

        # 4. Nhận diện Kiểm sát viên có chức danh KSV
        if re.search(r'^(?:ksv|kiểm\s*sát\s*viên)[\s.:\-_/]*', item_lower):
            c_name = clean_person_name(item)
            if c_name and c_name not in ksv_found:
                ksv_found.append(c_name)
            continue

        # 5. So khớp đích danh theo danh sách Lãnh đạo Viện
        matched_ldv = False
        for ldv in LANH_DAO_VIEN:
            if ldv.lower() in item_lower:
                if ldv not in ld_vien_found:
                    ld_vien_found.append(ldv)
                matched_ldv = True
                break
        if matched_ldv:
            continue

        # 6. So khớp đích danh theo danh sách Lãnh đạo Phòng
        matched_ldp = False
        for ldp in LANH_DAO_PHONG_2:
            if ldp.lower() in item_lower:
                if ldp not in ld_phong_found:
                    ld_phong_found.append(ldp)
                matched_ldp = True
                break
        if matched_ldp:
            continue

        # 7. Còn lại là Kiểm sát viên / Cán bộ giải quyết (Chỉ lấy khi là tên người hợp lệ)
        clean_name = clean_person_name(item)
        if clean_name and clean_name not in ksv_found:
            ksv_found.append(clean_name)

    cb_tl_clean = []
    if can_bo_thu_ly:
        for p in re.split(r'[,;|\n]+', str(can_bo_thu_ly).replace("\xa0", " ")):
            if is_procedural_status_text(p) or is_form_status_text(p):
                continue
            c_tl = clean_person_name(p)
            if c_tl and c_tl not in cb_tl_clean:
                cb_tl_clean.append(c_tl)
    for cv in cv_found:
        if cv not in cb_tl_clean:
            cb_tl_clean.append(cv)

    return (
        "; ".join(ld_vien_found),
        "; ".join(ld_phong_found),
        "; ".join(ksv_found),
        "; ".join(cb_tl_clean)
    )



COLUMNS_CASE_29 = [
    # Nhóm 1: Vụ án & Đơn vị địa bàn
    ("stt", "STT", 50),
    ("nguon_du_lieu", "Nguồn dữ liệu đối chiếu", 170),
    ("ma_vu_an", "Mã vụ án", 190),
    ("ten_vu_an", "Tên vụ án / Nội dung / QĐ TĐC", 260),
    ("tinh_thanh", "Tỉnh / Thành phố", 135),

    # Nhóm 2: Quyết định Khởi tố
    ("so_qd_khoi_to", "Số QĐ khởi tố", 105),
    ("ngay_qd_khoi_to", "Ngày ra QĐ khởi tố", 120),
    ("don_vi_khoi_to", "Đơn vị ra QĐ khởi tố", 210),

    # Nhóm 3: Thụ lý Kiểm sát
    ("so_thu_ly", "Số thụ lý", 95),
    ("ngay_thu_ly", "Ngày thụ lý", 110),
    ("loai_thu_ly", "Loại thụ lý", 130),
    ("vks_thu_ly", "Đơn vị VKS thụ lý", 240),

    # Nhóm 4: Tiến độ & Thời hiệu TNHS
    ("ngay_tiep_nhan", "Ngày tiếp nhận hồ sơ", 125),
    ("thoi_han_giai_quyet", "Thời hạn giải quyết", 125),
    ("ngay_het_thoi_hieu", "Ngày hết thời hiệu TNHS", 145),
    ("so_ngay_con_lai", "Số ngày còn lại", 115),
    ("muc_do_khan_cap", "Mức độ khẩn cấp", 175),
    ("canh_bao_thoi_hieu", "Cảnh báo thời hiệu", 180),

    # Nhóm 5: Bị can & Tội danh
    ("bi_can_dau_vu", "Bị can đầu vụ & CCCD", 200),
    ("tong_bi_can", "Tổng số bị can", 95),
    ("dieu_luat", "Thông tin tội danh / Điều luật", 250),
    ("loai_toi", "Phân loại tội", 130),

    # Nhóm 6: Phân công Lãnh đạo & Cán bộ nghiệp vụ (Tách riêng biệt)
    ("lanh_dao_vien", "Lãnh đạo Viện phụ trách", 180),
    ("lanh_dao_phong", "Lãnh đạo Phòng phụ trách", 210),
    ("can_bo_giai_quyet", "Kiểm sát viên giải quyết", 210),
    ("can_bo_thu_ly", "Cán bộ thụ lý / Chuyên viên", 170),

    # Nhóm 7: Tiến độ & Trạng thái tố tụng
    ("giai_doan", "Giai đoạn tố tụng", 115),
    ("trang_thai_vu_an", "Trạng thái vụ án", 150),
    ("trang_thai_bieu_mau", "Trạng thái biểu mẫu", 180),

    # Nhóm 8: Đơn vị VKS cập nhật mới nhất (Kế thừa từ rà soát thực tế)
    ("ten_don_vi_moi_nhat", "Tên đơn vị cập nhật mới nhất", 220),
]
COLUMNS_CASE_30 = COLUMNS_CASE_29
COLUMNS_CASE_27 = COLUMNS_CASE_29
COLUMNS_CASE_26 = COLUMNS_CASE_29
COLUMNS_CASE_24 = COLUMNS_CASE_29

# 30 Cột chuẩn hóa toàn diện cho Danh sách Tin báo về tội phạm (QLAHS)
COLUMNS_TIN_BAO_35 = [
    # Nhóm 1: Nhận diện & Nguồn tin
    ("stt", "STT", 50),
    ("ma_tin_bao", "Mã tin báo", 160),
    ("loai_tin", "Loại tin", 100),
    ("nguon_tin", "Phân loại nguồn tin", 120),
    ("tinh_tiet_phuc_tap", "Phức tạp", 80),

    # Nhóm 2: Người báo tin & Đối tượng bị tố giác
    ("nguoi_bao_tin", "Người / CQ báo tin", 180),
    ("nguoi_bi_to_giac", "Người / Đối tượng bị tố giác", 180),
    ("cccd_nguoi_bao", "CCCD / Định danh", 130),
    ("sdt_nguoi_bao", "Số điện thoại", 110),
    ("dia_chi_nguoi_bao", "Địa chỉ người báo", 200),

    # Nhóm 3: Nội dung tin báo & Tội danh
    ("thoi_gian_dia_diem", "Thời gian & Địa điểm", 180),
    ("noi_dung_tin_bao", "Nội dung tóm tắt nguồn tin", 300),
    ("dieu_luat", "Thông tin điều luật / Tội danh", 240),
    ("chuong_toi", "Chương BLHS", 180),
    ("muc_do_nghiem_trong", "Mức độ tội", 140),

    # Nhóm 4: Tiếp nhận & Thụ lý Kiểm sát
    ("don_vi_tiep_nhan", "Đơn vị tiếp nhận ban đầu", 220),
    ("ngay_tiep_nhan_tin", "Ngày tiếp nhận tin", 120),
    ("ngay_cqdt_nhan", "Ngày CQĐT thụ lý", 120),
    ("dieu_tra_vien", "Điều tra viên (ĐTV)", 150),
    ("ngay_vks_nhan", "Ngày VKS tiếp nhận", 120),
    ("so_thu_ly", "Số thụ lý (Số TL)", 100),
    ("ngay_vks_thu_ly", "Ngày VKS thụ lý", 120),
    ("vks_thu_ly", "Đơn vị VKS thụ lý", 200),

    # Nhóm 5: Cán bộ Kiểm sát thụ lý
    ("ksv_chinh", "Kiểm sát viên thụ lý chính", 180),
    ("ksv_phoi_hop", "KSV phối hợp / giúp việc", 160),
    ("don_vi_giai_quyet", "Đơn vị giải quyết", 180),

    # Nhóm 6: Thời hạn & Tiến độ giải quyết
    ("thoi_han_tu_ngay", "Thời hạn từ ngày", 110),
    ("thoi_han_den_ngay", "Thời hạn đến ngày", 110),
    ("so_ngay_con_lai", "Số ngày còn lại", 110),
    ("muc_do_khan_cap", "Cảnh báo tiến độ / Thời hạn", 180),
    ("ket_qua_giai_quyet", "Tình trạng / Kết quả giải quyết", 150),

    # Nhóm 7: Tạm đình chỉ & Thời hiệu truy cứu TNHS (Đặc biệt chú ý)
    ("ngay_tam_dinh_chi", "Ngày tạm đình chỉ", 110),
    ("ly_do_tam_dinh_chi", "Lý do tạm đình chỉ", 180),
    ("ngay_het_thoi_hieu", "Ngày hết thời hiệu truy cứu", 130),
    ("so_ngay_con_lai_thoi_hieu", "Số ngày còn lại thời hiệu", 130),
    ("canh_bao_thoi_hieu", "Cảnh báo thời hiệu / Khẩn cấp", 180),
    ("canh_bao_chi_tiet_thoi_hieu", "Cảnh báo chi tiết thời hiệu", 180),
    ("ghi_chu", "Ghi chú", 160),
]
COLUMNS_TIN_BAO_30 = COLUMNS_TIN_BAO_35
COLUMNS_TIN_BAO_RECONCILE_36 = [
    ("stt", "STT", 50),
    ("nguon_du_lieu", "Nguồn dữ liệu đối chiếu", 240),
] + [col for col in COLUMNS_TIN_BAO_35 if col[0] != "stt"]



# Pre-compile các mẫu nhận dạng 63 Tỉnh/Thành phố để tối ưu tốc độ x100 lần khi xử lý dữ liệu nặng
COMPILED_PROVINCE_RULES = []
for _pname, _patterns in PROVINCE_RULES:
    _escaped = [r"(?:\b|[\s,.\-_/])" + re.escape(p) + r"(?:\b|[\s,.\-_/]|$)" for p in _patterns]
    COMPILED_PROVINCE_RULES.append((_pname, re.compile("|".join(_escaped), flags=re.IGNORECASE)))

RE_FALLBACK_PROV = re.compile(r"(?:tỉnh|thành\s*phố|tp\.?|tx\.?)\s+([A-ZÀ-Ỹa-zà-ỹ\s]+?)(?:,|-|\.|$)", flags=re.IGNORECASE)
RE_GIAI_DOAN_START = re.compile(r"^giai\s*đoạn\s*:", flags=re.IGNORECASE)
RE_GIAI_DOAN_VAL = re.compile(r"^giai\s*đoạn\s*:\s*(.*)$", flags=re.IGNORECASE)
RE_NGAY_SINH = re.compile(r"^(?:ngày\s*sinh|năm\s*sinh|ns)\s*:\s*(.*)$", flags=re.IGNORECASE)
RE_CCCD = re.compile(r"^(?:cccd|số\s*cccd|cmnd|số\s*cmnd|đdcn|căn\s*cước[^:]*)\s*:\s*(.*)$", flags=re.IGNORECASE)


def detect_province(cqdt_text, fallback_text=""):
    """Đọc và nhận dạng Tỉnh/Thành phố từ chuỗi Cơ quan cảnh sát điều tra (Tối ưu tiền biên dịch)."""
    if not cqdt_text and not fallback_text:
        return "Chưa xác định"

    target = (str(cqdt_text or "") + " " + str(fallback_text or "")).lower()

    for prov_name, pattern_re in COMPILED_PROVINCE_RULES:
        if pattern_re.search(target):
            return prov_name

    # Dự phòng: Bắt các cụm từ "tỉnh X", "thành phố Y", "tp Z"
    m_tinh = RE_FALLBACK_PROV.search(target)
    if m_tinh:
        cand = m_tinh.group(1).strip().title()
        if len(cand) >= 2 and cand.lower() not in ("công an", "cảnh sát", "điều tra", "hồ sơ"):
            return cand

    return "Khác / Chưa rõ"


def clean_lines(cell_values):
    lines = []
    for val in cell_values:
        if val is None:
            continue
        for sub in str(val).replace("\r\n", "\n").replace("\r", "\n").split("\n"):
            s = sub.strip()
            if s and not s.startswith('="'):
                lines.append(s)
            elif s.startswith('="') and s.endswith('"'):
                lines.append(s[2:-1].strip())
    return lines


def parse_defendant_detail_rows(raw_rows, progress_cb=None, filename=""):
    """
    Bóc tách danh sách bị can chi tiết (định dạng 10 cột phẳng xuất từ QLAHS).
    Gồm: STT, Mã bị can, Họ tên, Ngày sinh, Giới tính, CCCD, Địa chỉ, Điều luật, Loại tội, Trạng thái.
    Hỗ trợ gom các dòng bổ sung khoản/điểm của cùng một bị can và chuẩn hóa ngày sinh DD/MM/YYYY.
    """
    if not raw_rows:
        return []

    header_rows = []
    for r_idx, row in enumerate(raw_rows):
        if not row:
            continue
        c0 = str(row[0] or "").strip().lower()
        c1 = str(row[1] or "").strip().lower() if len(row) > 1 and row[1] is not None else ""
        c2 = str(row[2] or "").strip().lower() if len(row) > 2 and row[2] is not None else ""
        if c0 == "stt" and ("mã bị can" in c1 or "họ tên" in c2):
            header_rows.append(r_idx)

    if not header_rows:
        header_rows = [0]

    records = []
    for blk_idx, hr in enumerate(header_rows):
        next_hr = header_rows[blk_idx + 1] if blk_idx + 1 < len(header_rows) else len(raw_rows)
        current_rec = None

        for r_idx in range(hr + 1, next_hr):
            row = raw_rows[r_idx]
            if not row:
                continue
            c0 = str(row[0] or "").strip().lower()
            c1 = str(row[1] or "").strip() if len(row) > 1 and row[1] is not None else ""
            c2 = str(row[2] or "").strip() if len(row) > 2 and row[2] is not None else ""
            c7 = str(row[7] or "").strip() if len(row) > 7 and row[7] is not None else ""

            if c0 == "stt" and ("mã bị can" in c1.lower() or "họ tên" in c2.lower()):
                continue

            if not c1 and not c2 and c7 and current_rec:
                current_rec["dieu_luat"] += "\n" + c7
                continue

            if c1 and (c1.startswith("VA_") or c1.startswith("E01") or c1.startswith("010") or re.match(r"^[A-Za-z0-9_\.]+$", c1)):
                if current_rec:
                    records.append(current_rec)

                ns_val = row[3] if len(row) > 3 else None
                ns_str = ""
                if isinstance(ns_val, (datetime, date)):
                    ns_str = ns_val.strftime("%d/%m/%Y")
                elif ns_val:
                    ns_raw = str(ns_val).strip()[:10]
                    m_dt = re.match(r"^(\d{4})-(\d{2})-(\d{2})", ns_raw)
                    if m_dt:
                        ns_str = f"{m_dt.group(3)}/{m_dt.group(2)}/{m_dt.group(1)}"
                    else:
                        ns_str = ns_raw

                cccd_val = str(row[5] or "").strip() if len(row) > 5 and row[5] is not None else ""
                if cccd_val.isdigit() and len(cccd_val) == 11:
                    cccd_val = "0" + cccd_val

                m_va = ""
                m_match = re.match(r"^(VA_[A-Za-z0-9_\.]+?)_\d{2,3}$", c1)
                if m_match:
                    m_va = m_match.group(1)

                addr = str(row[6] or "").strip() if len(row) > 6 and row[6] is not None else ""
                tinh_val = detect_province(addr) if addr else ""

                current_rec = {
                    "stt": len(records) + 1,
                    "nguon_du_lieu": filename,
                    "ma_vu_an": m_va,
                    "ten_vu_an": "",
                    "cqdt": "",
                    "tinh_thanh": tinh_val,
                    "giai_doan": "Điều tra",
                    "ma_bi_can": c1,
                    "ten_bi_can": c2,
                    "ngay_sinh": ns_str,
                    "so_cccd": cccd_val,
                    "gioi_tinh": str(row[4] or "").strip() if len(row) > 4 and row[4] is not None else "",
                    "dia_chi": addr,
                    "dieu_luat": c7,
                    "loai_toi": str(row[8] or "").strip() if len(row) > 8 and row[8] is not None else "",
                    "trang_thai": str(row[9] or "").strip() if len(row) > 9 and row[9] is not None else "",
                    "block_idx": blk_idx
                }

        if current_rec:
            records.append(current_rec)

    return records


def parse_qlahs_rows(raw_rows, progress_cb=None, filename=""):
    """Bóc tách danh sách dòng thô từ Excel QLAHS thành danh sách bản ghi 12 cột chuẩn.
    Tự động hỗ trợ cả 2 định dạng:
    - Định dạng A: Khối 5 cột thô (Bị can, Đình nã, Truy nã)
    - Định dạng B: Danh sách 10 cột chi tiết (Mã bị can, Họ tên, CCCD, Địa chỉ...)
    """
    if not raw_rows:
        return []

    is_detail_format = any(
        isinstance(row, (list, tuple)) and len(row) > 1 and
        str(row[0] or "").strip().lower() == "stt" and
        "mã bị can" in str(row[1] or "").strip().lower()
        for row in raw_rows[:15]
    )
    if is_detail_format:
        return parse_defendant_detail_rows(raw_rows, progress_cb, filename=filename)

    blocks = []
    current_block = None

    for row in raw_rows:
        vals = [row[i] if i < len(row) else None for i in range(5)]
        texts = []
        for v in vals:
            if v is None:
                texts.append("")
            else:
                s = str(v).strip()
                if s.startswith('="') and s.endswith('"'):
                    s = s[2:-1].strip()
                texts.append(s)

        if not any(texts):
            continue

        c0_low = texts[0].lower()
        c1_low = texts[1].lower()
        c2_low = texts[2].lower()
        if c0_low == "stt" or "thông tin hồ sơ" in c1_low or (c1_low == "" and c2_low == "họ tên"):
            continue

        b_lines = [x.strip() for x in texts[1].replace("\r\n", "\n").split("\n") if x.strip()]
        starts_new_block = any(RE_GIAI_DOAN_START.search(ln) for ln in b_lines)

        if starts_new_block:
            if current_block is not None:
                blocks.append(current_block)
            current_block = {"A": [], "B": [], "C": [], "D": [], "E": []}

        if current_block is None:
            current_block = {"A": [], "B": [], "C": [], "D": [], "E": []}

        for idx, key in enumerate(["A", "B", "C", "D", "E"]):
            if texts[idx]:
                current_block[key].append(texts[idx])

    if current_block is not None and any(current_block.values()):
        blocks.append(current_block)

    total_blocks = len(blocks)
    records = []
    for idx, blk in enumerate(blocks, start=1):
        if progress_cb and total_blocks > 300 and idx % 200 == 0:
            progress_cb(
                f"Đang chuẩn hóa & nhận dạng Tỉnh/Thành: {idx:,}/{total_blocks:,} bị can...",
                "Đang phân tích các giai đoạn và điều luật..."
            )

        a_lines = clean_lines(blk["A"])
        b_lines = clean_lines(blk["B"])
        c_lines = clean_lines(blk["C"])
        d_lines = clean_lines(blk["D"])
        e_lines = clean_lines(blk["E"])

        stt_val = idx
        for a_ln in a_lines:
            if a_ln.isdigit():
                stt_val = int(a_ln)
                break

        giai_doan = ""
        b_rem = []
        for ln in b_lines:
            m_gd = RE_GIAI_DOAN_VAL.match(ln)
            if m_gd and not giai_doan:
                giai_doan = m_gd.group(1).strip()
            else:
                b_rem.append(ln)

        ma_vu_an, ten_vu_an, cqdt = "", "", ""
        if len(b_rem) >= 3:
            ma_vu_an = b_rem[0]
            cqdt = b_rem[-1]
            ten_vu_an = " - ".join(b_rem[1:-1])
        elif len(b_rem) == 2:
            ma_vu_an = b_rem[0]
            if any(b_rem[1].lower().startswith(p) for p in CQDT_PREFIXES):
                cqdt = b_rem[1]
            else:
                ten_vu_an = b_rem[1]
        elif len(b_rem) == 1:
            ma_vu_an = b_rem[0]

        # Tự động đọc và nhận dạng Tỉnh/Thành bằng regex tiền biên dịch siêu tốc
        tinh_thanh = detect_province(cqdt, ten_vu_an)

        ngay_sinh, so_cccd = "", ""
        c_rem = []
        for ln in c_lines:
            m_ns = RE_NGAY_SINH.match(ln)
            m_cccd = RE_CCCD.match(ln)
            if m_ns and not ngay_sinh:
                ngay_sinh = m_ns.group(1).strip()
            elif m_cccd and not so_cccd:
                so_cccd = m_cccd.group(1).strip()
            else:
                c_rem.append(ln)

        ma_bi_can, ten_bi_can = "", ""
        if len(c_rem) >= 2:
            ma_bi_can = c_rem[0]
            ten_bi_can = " ".join(c_rem[1:])
        elif len(c_rem) == 1:
            if re.search(r"\d", c_rem[0]):
                ma_bi_can = c_rem[0]
            else:
                ten_bi_can = c_rem[0]

        dieu_luat = "; ".join(dict.fromkeys(d_lines))
        trang_thai = "; ".join(dict.fromkeys(e_lines))

        records.append({
            "stt": stt_val,
            "ma_vu_an": ma_vu_an,
            "ten_vu_an": ten_vu_an,
            "cqdt": cqdt,
            "tinh_thanh": tinh_thanh,
            "giai_doan": giai_doan,
            "ma_bi_can": ma_bi_can,
            "ten_bi_can": ten_bi_can,
            "ngay_sinh": ngay_sinh,
            "so_cccd": so_cccd,
            "dieu_luat": dieu_luat,
            "trang_thai": trang_thai,
        })

    return records


def extract_officers_from_text(text):
    """Trích xuất KSV, PVT, CV từ chuỗi văn bản tên vụ án (vd: (PVT Tùng, KSV Đông))."""
    if not text:
        return ""
    m = re.search(r'\(([^)]*(?:ksv|pvt|cv|tp|ptp)[^)]*)\)', text, re.IGNORECASE)
    if m:
        parts = re.split(r'[,;|\-]', m.group(1).strip())
        return "; ".join(p.strip() for p in parts if p.strip())
    return ""

def clean_vks_unit_name(u):
    """Rút gọn và chuẩn hóa tên đơn vị VKS cho báo cáo và bảng tổng hợp:
    - Phòng THQCT, KSĐT, KSXXST án trật tự xã hội (P2) -> Phòng 2 - VKSND TP.HCM
    - Phòng Công tố và Kiểm sát điều tra, Kiểm sát xét xử sơ thẩm án trật tự xã hội (Phòng 2) -> Phòng 2 - VKSND TP.HCM
    - Viện Kiểm sát nhân dân Khu vực 4 - Hồ Chí Minh -> VKSND Khu vực 4
    """
    if not u:
        return "Chưa rõ đơn vị VKS"
    u_str = str(u).replace("\\", "/").strip()
    parts = [p.strip() for p in u_str.split("/") if p.strip()]
    first = parts[0] if parts else u_str
    u_low = u_str.lower()

    # 0. Kiểm tra từ điển viết tắt tùy biến theo cấu hình người dùng (Settings)
    custom_map = CURRENT_CONFIG.get("unit_aliases", {}) if "CURRENT_CONFIG" in globals() else {}
    for alias_key, target_name in custom_map.items():
        if alias_key.lower() == first.lower() or alias_key.lower() == u_low:
            return target_name
        if f"({alias_key.lower()})" in u_low:
            return target_name
        if re.search(r'\b' + re.escape(alias_key.lower()) + r'\b', u_low):
            return target_name

    # 1. Nhận diện các Phòng nghiệp vụ VKSND TP.HCM
    # 1.1. Có ký hiệu viết tắt trong ngoặc: (P2), (P.2), (P 2), (Phòng 2), (Phòng. 2)
    m_p_paren = re.search(r'\((?:phòng|p)\.?\s*(\d+)\)', u_str, re.IGNORECASE)
    if m_p_paren:
        return f"Phòng {int(m_p_paren.group(1))} - VKSND TP.HCM"

    # 1.2. Nhận diện theo tên chuyên môn đặc thù của các Phòng VKSND TP.HCM
    if "trật tự xã hội" in u_low and any(k in u_low for k in ["phòng", "thqct", "công tố", "ksđt", "ksxxst", "p2", "vks"]):
        return "Phòng 2 - VKSND TP.HCM"
    if ("an ninh" in u_low or "ma túy" in u_low) and any(k in u_low for k in ["phòng", "thqct", "công tố", "ksđt", "ksxxst", "p1", "vks"]):
        return "Phòng 1 - VKSND TP.HCM"
    if ("kinh tế" in u_low or "tham nhũng" in u_low or "chức vụ" in u_low) and any(k in u_low for k in ["phòng", "thqct", "công tố", "ksđt", "ksxxst", "p3", "vks"]):
        return "Phòng 3 - VKSND TP.HCM"

    # 1.3. Nhận diện dạng Phòng X hoặc PX khi có thông tin VKS / TP.HCM
    m_phong_chu = re.search(r'\b(?:phòng|p)\.?\s*(\d+)\b', first, re.IGNORECASE)
    if m_phong_chu and any(k in u_low for k in ["hồ chí minh", "thành phố", "vks", "kiểm sát", "thqct"]):
        return f"Phòng {int(m_phong_chu.group(1))} - VKSND TP.HCM"

    # 1.4. Đã là Phòng X - VKSND TP.HCM chuẩn
    m_std = re.search(r'phòng\s*(\d+)\s*-\s*vksnd\s*(?:tp\.?\s*hcm|thành\s*phố\s*hồ\s*chí\s*minh)', u_low)
    if m_std:
        return f"Phòng {int(m_std.group(1))} - VKSND TP.HCM"

    # 1.5. Nhận diện dạng Phòng X hoặc PX đơn lẻ (ví dụ: PHÒNG 2, PHÒNG 7, PHÒNG 3)
    m_p_solo = re.search(r'^(?:phòng|p)\.?\s*([1-9]|1[0-5])$', u_str.strip(), re.IGNORECASE)
    if m_p_solo:
        return f"Phòng {int(m_p_solo.group(1))} - VKSND TP.HCM"

    # 2. Rút gọn tên VKS Khu vực
    m_kv = re.search(r'Khu\s*vực\s*0*([1-9]|1[0-9])\b', u_str, re.IGNORECASE)
    if m_kv:
        return f"VKSND Khu vực {int(m_kv.group(1))}"

    # 3. Rút gọn tên VKS cấp Quận/Huyện cũ
    m_qh = re.search(r'VKSND\s+(Quận|Huyện|Thành phố|TX)\s+([^/-]+)', first, re.IGNORECASE)
    if m_qh:
        return f"VKSND {m_qh.group(1)} {m_qh.group(2).strip()}"

    if "viện kiểm sát nhân dân thành phố hồ chí minh" in u_low or "vksnd thành phố hồ chí minh" in u_low or "vksnd tp.hcm" in u_low:
        return "VKSND Thành phố Hồ Chí Minh"

    if any(first.lower().endswith(ext) for ext in [".xlsx", ".xls", ".csv"]):
        return "VKSND Thành phố Hồ Chí Minh"

    return first


def extract_unit_from_source(source_text):
    """
    Trích xuất chính xác tên đơn vị VKS thụ lý từ tên file hoặc nguồn dữ liệu đối chiếu.
    Ví dụ:
    - 'KHU VỰC 6.ĐT.xlsx' -> 'VKSND Khu vực 6'
    - 'KHU VỰC 16.TT.xlsx' -> 'VKSND Khu vực 16'
    - 'Khu vực 1.ST.xlsx' -> 'VKSND Khu vực 1'
    - 'KV04_DT.xlsx' -> 'VKSND Khu vực 4'
    - 'Phòng 2.xlsx', 'P2.xlsx' -> 'Phòng 2 - VKSND TP.HCM'
    """
    if not source_text:
        return None
    s = str(source_text).strip()
    s_low = s.lower()

    # 0. Kiểm tra cấu hình unit_aliases từ người dùng (Settings)
    custom_map = CURRENT_CONFIG.get("unit_aliases", {}) if "CURRENT_CONFIG" in globals() else {}
    for alias_key, target_name in custom_map.items():
        if alias_key.lower() in s_low:
            return target_name

    # 1. Tìm mẫu Khu vực (1 đến 19)
    # Hỗ trợ: 'khu vực 6', 'khu vuc 6', 'khuvuc 6', 'kv6', 'kv 6', 'kv06', 'k.v 6', 'k.vực 6'
    m_kv = re.search(r'(?:khu[_\s\-]*(?:v[ưựuứừửữ]c|vuc)?|kv|k\.v)[_\s\-]*0*([1-9]|1[0-9])(?!\d)', s_low, re.IGNORECASE)
    if m_kv:
        kv_num = int(m_kv.group(1))
        if 1 <= kv_num <= 19:
            return f"VKSND Khu vực {kv_num}"

    # 2. Tìm mẫu Phòng ban VKSND TP.HCM
    if any(k in s_low for k in ["phòng 2", "phong_2", "phong 2", "p2", "p.2", "trật tự xã hội", "trat_tu_xa_hoi", "ttxh"]):
        return "Phòng 2 - VKSND TP.HCM"
    if any(k in s_low for k in ["phòng 1", "phong_1", "phong 1", "p1", "p.1", "an ninh", "ma túy", "an_ninh", "ma_tuy"]):
        return "Phòng 1 - VKSND TP.HCM"
    if any(k in s_low for k in ["phòng 3", "phong_3", "phong 3", "p3", "p.3", "kinh tế", "tham nhũng", "kinh_te", "tham_nhung"]):
        return "Phòng 3 - VKSND TP.HCM"
    m_p = re.search(r'(?:phòng|phong|p)[_\s\-]*0*([1-9]|1[0-9])(?!\d)', s_low, re.IGNORECASE)
    if m_p:
        p_num = int(m_p.group(1))
        return f"Phòng {p_num} - VKSND TP.HCM"

    # 3. Tìm các quận huyện cũ TP.HCM nếu có trong tên file
    for q_num in range(1, 13):
        if re.search(rf'\bqu[ậa]n[_\s\-]*0*{q_num}\b', s_low, re.I):
            return f"VKSND Quận {q_num}"

    return None


def detect_file_unit(filename, raw_rows):
    """Xác định đơn vị VKS phụ trách tệp dữ liệu dựa vào tên file và các mã mới trong file."""
    if filename:
        u_src = extract_unit_from_source(filename)
        if u_src:
            return u_src

    # Quét các mã mới có trong file để tìm đơn vị chủ đạo (chỉ khi 100% đồng nhất từ một đơn vị)
    kv_counts = {}
    tp_count = 0
    for row in raw_rows:
        row_str = " ".join(str(c or "") for c in row)
        for m in re.finditer(r'VA_E01\.59[._]0*([1-9]|1[0-9])_(\d{4})_(\d+)', row_str, re.I):
            kv = int(m.group(1))
            kv_counts[kv] = kv_counts.get(kv, 0) + 1
        for m in re.finditer(r'VA_E01\.59_(\d{4})_(\d+)', row_str, re.I):
            tp_count += 1

    if kv_counts and len(kv_counts) == 1 and tp_count == 0 and list(kv_counts.values())[0] >= 3:
        top_kv = list(kv_counts.keys())[0]
        return f"VKSND Khu vực {top_kv}"
    if tp_count >= 3 and not kv_counts:
        return "VKSND Thành phố Hồ Chí Minh"
    return None


def parse_case_code_unit(case_code, file_context_unit=None):
    """
    Phân tích mã vụ án theo cấu trúc chuẩn QLAHS:
    1. Mã mới có dấu gạch dưới _:
       - Cấp Thành phố (các phòng nghiệp vụ):
         VA_E01.59_YYMM_STT (ví dụ VA_E01.59_2610_0158) -> Không có cụm mã khu vực 3 chữ số
         -> Trả về ("TP. Hồ Chí Minh", "VKSND Thành phố Hồ Chí Minh")
       - Cấp Khu vực (Khu vực 1 -> 19):
         VA_E01.59_XXX_YYMM_STT hoặc VA_E01.59.XXX_YYMM_STT (XXX: 001 -> 019)
         -> Trả về ("TP. Hồ Chí Minh", f"VKSND Khu vực {int(XXX)}")
    2. Mã cũ (dùng dấu chấm ., hoặc không có định dạng mã mới):
       - Tuyệt đối không suy đoán số lung tung từ trong mã cũ!
       - Ưu tiên số 1: Lấy trực tiếp theo file_context_unit (tên file hoặc nguồn đối chiếu) nếu có.
       - Nếu không có file_context_unit, mới để 'Chưa xác định đơn vị (Mã cũ QLAHS)'.
    """
    if not case_code:
        return "TP. Hồ Chí Minh", clean_vks_unit_name(file_context_unit) if file_context_unit else "Chưa xác định đơn vị"

    code = str(case_code).strip()

    # 1.1. Mã mới VKS Khu vực
    m_kv = re.search(r'VA_E01\.59[._]0*([1-9]|1[0-9])_(\d{4})_(\d+)', code, re.I)
    if m_kv:
        kv_num = int(m_kv.group(1))
        if 1 <= kv_num <= 19:
            return "TP. Hồ Chí Minh", f"VKSND Khu vực {kv_num}"

    # 1.2. Mã mới VKS Thành phố (các phòng nghiệp vụ)
    m_tp = re.search(r'VA_E01\.59_(\d{4})_(\d+)', code, re.I)
    if m_tp:
        if file_context_unit and ("phòng" in file_context_unit.lower() or "p2" in file_context_unit.lower()):
            return "TP. Hồ Chí Minh", clean_vks_unit_name(file_context_unit)
        return "TP. Hồ Chí Minh", "VKSND Thành phố Hồ Chí Minh"

    # 2. Mã cũ (tất cả các mã khác: VA_010..., VA_E01.59.XX...):
    # Lấy luôn tên đơn vị từ nguồn đối chiếu / tên file nếu có!
    if file_context_unit:
        return "TP. Hồ Chí Minh", clean_vks_unit_name(file_context_unit)
    return "TP. Hồ Chí Minh", "Chưa xác định đơn vị (Mã cũ QLAHS)"


def parse_tin_bao_code_unit(tb_code, file_context_unit=None):
    """
    Tự động giải mã đơn vị VKS thụ lý trực tiếp từ mã Tin báo tội phạm:
    1. Mã mới (có tiền tố TB_ hoặc TB- phía trước):
       - TB_E01.59.015_... -> VKSND Khu vực 15
       - TB_E01.59_... -> VKSND Thành phố Hồ Chí Minh (hoặc file_context_unit)
    2. Mã cũ (KHÔNG có TB_ phía trước, ví dụ: E01.59.32 - ..., 010203 - ...):
       - Do không có quy luật cố định, nếu không có nguồn đối chiếu / file_context_unit
         thì để 'Chưa xác định đơn vị (Mã cũ QLAHS)'.
    """
    valid_context = None
    if file_context_unit:
        u_cand = clean_vks_unit_name(file_context_unit)
        if u_cand and "chưa" not in u_cand.lower() and not any(u_cand.lower().endswith(ext) for ext in [".xlsx", ".xls", ".csv"]):
            valid_context = u_cand

    if not tb_code:
        return valid_context or "Chưa xác định đơn vị (Mã cũ QLAHS)"

    code = str(tb_code).strip()
    is_new_tb_code = code.upper().startswith("TB_") or code.upper().startswith("TB-")

    if is_new_tb_code:
        # 1. Mã mới: Có quy luật xác định đơn vị chuẩn
        m_kv = re.search(r'TB[_\-]E01\.59[._]0*([1-9]|1[0-9])(?=[._\s\-]|\b)', code, re.I)
        if m_kv:
            kv_num = int(m_kv.group(1))
            if 1 <= kv_num <= 19:
                return f"VKSND Khu vực {kv_num}"
        # Cấp Thành phố (các phòng nghiệp vụ hoặc VKSND TP.HCM)
        if re.search(r'TB[_\-]E01\.59', code, re.I):
            return valid_context or "VKSND Thành phố Hồ Chí Minh"
        return valid_context or "VKSND Thành phố Hồ Chí Minh"

    # 2. Mã cũ (KHÔNG có TB_ phía trước):
    # Không có quy luật suy đoán -> Nếu có nguồn đơn vị từ file thì lấy, ngược lại để Chưa xác định đơn vị (Mã cũ QLAHS)
    return valid_context or "Chưa xác định đơn vị (Mã cũ QLAHS)"


def detect_district_from_case_code(case_code, default_unit=None):
    """Alias tương thích ngược."""
    return parse_case_code_unit(case_code, default_unit)


def detect_procedural_stage(status_text="", filename="", loai_thu_ly="", current_stage=""):
    """
    Tự động nhận diện và phân loại chính xác tuyệt đối Giai đoạn tố tụng của vụ án hình sự:
    - Sơ thẩm (Xét xử sơ thẩm)
    - Truy tố
    - Điều tra (Khởi tố - Điều tra)
    - Phúc thẩm (Xét xử phúc thẩm)
    - Thi hành án
    - GĐT/TT (Giám đốc thẩm / Tái thẩm)
    - Tin báo (Giai đoạn tiền khởi tố)

    Quy tắc xét theo thứ tự ưu tiên chặt chẽ:
    1. Tiền tố có dấu hai chấm ':' trong Trạng thái vụ án (Chuẩn tuyệt đối từ QLAHS)
    2. Từ khóa đặc thù trong Trạng thái vụ án
    3. Từ khóa trong Loại thụ lý
    4. Ký hiệu phân loại giai đoạn trong Tên file (ví dụ: .ĐT., .TT., .ST., .PT., .THA.)
    5. Giữ lại giá trị giai đoạn hiện tại nếu đã hợp lệ
    6. Mặc định 'Điều tra' cho hồ sơ vụ án nếu chưa đủ thông tin
    """
    st_low = str(status_text or "").strip().lower()
    fn_low = str(filename or "").strip().lower()
    lt_low = str(loai_thu_ly or "").strip().lower()
    cur_low = str(current_stage or "").strip().lower()

    # 1. Kiểm tra tiền tố có dấu hai chấm ':' trong Trạng thái vụ án (Chuẩn tuyệt đối QLAHS)
    if ":" in st_low:
        prefix = st_low.split(":", 1)[0].strip()
        if any(k in prefix for k in ("sơ thẩm", "st", "xét xử sơ thẩm", "toà án sơ thẩm")):
            return "Sơ thẩm"
        if any(k in prefix for k in ("truy tố", "tt", "giai đoạn truy tố")):
            return "Truy tố"
        if any(k in prefix for k in ("điều tra", "đt", "dt", "ksđt", "kiểm sát điều tra")):
            return "Điều tra"
        if any(k in prefix for k in ("phúc thẩm", "pt", "xét xử phúc thẩm")):
            return "Phúc thẩm"
        if any(k in prefix for k in ("thi hành án", "tha")):
            return "Thi hành án"
        if any(k in prefix for k in ("giám đốc thẩm", "tái thẩm", "gđt", "tt")):
            return "GĐT/TT"
        if any(k in prefix for k in ("tin báo", "tố giác", "tiền khởi tố")):
            return "Tin báo"

    # 2. Kiểm tra từ khóa trong Trạng thái vụ án (khi không có dấu hai chấm)
    if any(k in st_low for k in ("sơ thẩm", "sau xét xử", "tuyên án", "hội đồng xét xử", "chuẩn bị xét xử sơ thẩm")):
        return "Sơ thẩm"
    if any(k in st_low for k in ("truy tố", "cáo trạng", "qđ truy tố", "quyết định truy tố")):
        return "Truy tố"
    if any(k in st_low for k in ("điều tra", "ksđt", "kiểm sát điều tra", "kết luận điều tra", "thqct và kiểm sát điều tra", "phục hồi điều tra")):
        return "Điều tra"
    if any(k in st_low for k in ("phúc thẩm", "kháng cáo", "kháng nghị", "y án", "hủy án", "sửa án")):
        return "Phúc thẩm"
    if any(k in st_low for k in ("thi hành án", "ủy thác thi hành án")):
        return "Thi hành án"

    # 3. Kiểm tra từ Loại thụ lý
    if "sơ thẩm" in lt_low or "xét xử sơ thẩm" in lt_low:
        return "Sơ thẩm"
    if "truy tố" in lt_low:
        return "Truy tố"
    if "điều tra" in lt_low or "ksđt" in lt_low:
        return "Điều tra"
    if "phúc thẩm" in lt_low or "xét xử phúc thẩm" in lt_low:
        return "Phúc thẩm"

    # 4. Kiểm tra ký hiệu giai đoạn từ Tên file (Đặc thù phân khu vực của VKS: KHU VỰC 6.ĐT.xlsx, KHU VỰC 7.TT.xlsx, KHU VỰC 19.ST.xlsx...)
    if re.search(r'[\._\-\s](st|so[\s_]*tham)[\._\-\s]', fn_low) or fn_low.endswith('.st.xlsx') or fn_low.endswith('.st.xls'):
        return "Sơ thẩm"
    if re.search(r'[\._\-\s](tt|truy[\s_]*to)[\._\-\s]', fn_low) or fn_low.endswith('.tt.xlsx') or fn_low.endswith('.tt.xls'):
        return "Truy tố"
    if re.search(r'[\._\-\s](đt|dt|dieu[\s_]*tra)[\._\-\s]', fn_low) or fn_low.endswith('.dt.xlsx') or fn_low.endswith('.đt.xlsx'):
        return "Điều tra"
    if re.search(r'[\._\-\s](pt|phuc[\s_]*tham)[\._\-\s]', fn_low):
        return "Phúc thẩm"
    if re.search(r'[\._\-\s](tha|thi[\s_]*hanh[\s_]*an)[\._\-\s]', fn_low):
        return "Thi hành án"

    # 5. Giữ lại giá trị giai đoạn hiện tại nếu đã hợp lệ
    for valid_s in ("Sơ thẩm", "Truy tố", "Điều tra", "Phúc thẩm", "Thi hành án", "GĐT/TT", "Tin báo"):
        if valid_s.lower() == cur_low:
            return valid_s

    return "Điều tra"


def update_case_urgency_and_statute(rec, filename="", today_date=None):
    """
    Tính toán và đồng bộ Số ngày còn lại (so_ngay_con_lai), Cảnh báo tiến độ / Khẩn cấp (muc_do_khan_cap),
    và Cảnh báo chi tiết (canh_bao_thoi_hieu) dựa trên Ngày hết thời hiệu và các cảnh báo nguồn.
    Quy tắc nghiệp vụ:
    - Nếu có ngày hết thời hiệu truy cứu TNHS: tính số ngày còn lại so với ngày hiện tại (today).
      + days_left < 0: 'Đã quá X ngày' -> '⚠️ Đã hết thời hiệu truy cứu TNHS (Quá X ngày)'
      + days_left == 0: 'Hôm nay hết hạn' -> '🔴 Hết thời hiệu hôm nay (0 ngày)'
      + 1 <= days_left <= 7: 'Còn X ngày' -> '🔴 Sắp hết thời hiệu truy cứu TNHS (Còn X ngày)'
      + 8 <= days_left <= 15: 'Còn X ngày' -> '🟠 Sắp hết thời hiệu truy cứu TNHS (Còn X ngày)'
      + 16 <= days_left <= 30: 'Còn X ngày' -> '🟡 Sắp hết thời hiệu truy cứu TNHS (Còn X ngày)'
      + days_left > 30: 'Còn X ngày' -> '🟢 Trong thời hiệu truy cứu TNHS (Còn X ngày)'
    - Nếu không có ngày cụ thể nhưng văn bản có cảnh báo:
      + 'hết thời hiệu' / 'đã hết': '⚠️ Đã hết thời hiệu truy cứu TNHS'
      + 'sắp hết': '🔴 Sắp hết thời hiệu truy cứu TNHS'
      + 'quá hạn': '⚠️ Quá hạn giải quyết'
    Trả về True nếu vụ án thuộc diện cảnh báo thời hiệu / quá hạn (để không bị trạng thái tạm đình chỉ ghi đè).
    """
    if today_date is None:
        today_date = date.today()

    th_val = str(rec.get("ngay_het_thoi_hieu") or "").strip()
    d_het_thoi_hieu = parse_date_safely(th_val)

    cb_raw = str(rec.get("canh_bao_thoi_hieu") or rec.get("canh_bao_qua_han") or "").strip()
    fn_low = str(filename or rec.get("nguon_du_lieu") or "").lower()

    # 1. Trường hợp có Ngày hết thời hiệu truy cứu TNHS cụ thể -> Tính toán số ngày còn lại
    if d_het_thoi_hieu:
        days_left = (d_het_thoi_hieu - today_date).days
        if days_left < 0:
            qua_ngay = -days_left
            rec["so_ngay_con_lai"] = f"Đã quá {qua_ngay} ngày"
            rec["muc_do_khan_cap"] = f"⚠️ Đã hết thời hiệu truy cứu TNHS (Quá {qua_ngay} ngày)"
            rec["canh_bao_thoi_hieu"] = f"Đã hết thời hiệu truy cứu TNHS (Quá {qua_ngay} ngày)"
        elif days_left == 0:
            rec["so_ngay_con_lai"] = "Hôm nay hết hạn"
            rec["muc_do_khan_cap"] = "🔴 Hết thời hiệu hôm nay (0 ngày)"
            rec["canh_bao_thoi_hieu"] = "Hết thời hiệu hôm nay (0 ngày)"
        elif days_left <= 7:
            rec["so_ngay_con_lai"] = f"Còn {days_left} ngày"
            rec["muc_do_khan_cap"] = f"🔴 Sắp hết thời hiệu truy cứu TNHS (Còn {days_left} ngày)"
            rec["canh_bao_thoi_hieu"] = f"Sắp hết thời hiệu truy cứu TNHS (Còn {days_left} ngày)"
        elif days_left <= 15:
            rec["so_ngay_con_lai"] = f"Còn {days_left} ngày"
            rec["muc_do_khan_cap"] = f"🟠 Sắp hết thời hiệu truy cứu TNHS (Còn {days_left} ngày)"
            rec["canh_bao_thoi_hieu"] = f"Sắp hết thời hiệu truy cứu TNHS (Còn {days_left} ngày)"
        elif days_left <= 30:
            rec["so_ngay_con_lai"] = f"Còn {days_left} ngày"
            rec["muc_do_khan_cap"] = f"🟡 Sắp hết thời hiệu truy cứu TNHS (Còn {days_left} ngày)"
            rec["canh_bao_thoi_hieu"] = f"Sắp hết thời hiệu truy cứu TNHS (Còn {days_left} ngày)"
        else:
            rec["so_ngay_con_lai"] = f"Còn {days_left} ngày"
            rec["muc_do_khan_cap"] = f"🟢 Trong thời hiệu truy cứu TNHS (Còn {days_left} ngày)"
            rec["canh_bao_thoi_hieu"] = f"Trong thời hiệu truy cứu TNHS (Còn {days_left} ngày)"
        return True

    # 2. Không có ngày cụ thể, nhưng có văn bản cảnh báo thời hiệu / quá hạn từ nội dung vụ án hoặc tên file
    cb_low = cb_raw.lower()
    has_cb = any(x in cb_low for x in ("hết thời hiệu", "đã hết", "sắp hết", "quá hạn", "quá thời hiệu"))
    has_fn_cb = ("sắp hết thời hiệu" in fn_low or "hết thời hiệu" in fn_low)

    if has_cb or has_fn_cb:
        if "đã hết" in cb_low or ("hết thời hiệu" in cb_low and "sắp" not in cb_low) or ("hết thời hiệu" in fn_low and "sắp" not in fn_low):
            rec["so_ngay_con_lai"] = rec.get("so_ngay_con_lai") or "Đã hết thời hiệu"
            rec["muc_do_khan_cap"] = "⚠️ Đã hết thời hiệu truy cứu TNHS"
            rec["canh_bao_thoi_hieu"] = cb_raw if cb_raw and "tạm đình chỉ" not in cb_low else "Đã hết thời hiệu truy cứu TNHS"
            return True
        elif "sắp hết" in cb_low or "sắp hết" in fn_low:
            rec["so_ngay_con_lai"] = rec.get("so_ngay_con_lai") or "Sắp hết thời hiệu"
            rec["muc_do_khan_cap"] = "🔴 Sắp hết thời hiệu truy cứu TNHS"
            rec["canh_bao_thoi_hieu"] = cb_raw if cb_raw and "tạm đình chỉ" not in cb_low else "Sắp hết thời hiệu truy cứu TNHS"
            return True
        elif "quá hạn" in cb_low:
            rec["so_ngay_con_lai"] = rec.get("so_ngay_con_lai") or "Quá hạn"
            rec["muc_do_khan_cap"] = cb_raw if cb_raw.startswith("⚠️") else f"⚠️ {cb_raw}"
            rec["canh_bao_thoi_hieu"] = cb_raw
            return True

    return False


def normalize_tin_bao_status(rec):
    """
    Chuẩn hóa trạng thái và cảnh báo tiến độ cho nguồn tin báo, tố giác tội phạm:
    Theo Điều 148 BLTTHS 2015, giai đoạn này là tiền khởi tố, do đó:
    - Khi tạm đình chỉ -> '⏸️ Tạm đình chỉ giải quyết' (hoặc 'Tạm đình chỉ giải quyết nguồn tin'),
      TUYỆT ĐỐI KHÔNG dùng 'Tạm đình chỉ vụ án' vì chưa hề có quyết định khởi tố vụ án.
    - Cảnh báo thời hiệu truy cứu TNHS theo Đ.230 BLTTHS & Đ.27 BLHS.
    """
    if not rec:
        return

    today_date = date.today()
    kq_raw = str(rec.get("ket_qua_giai_quyet") or "").strip()
    th_raw = rec.get("ngay_het_thoi_hieu")
    d_th = parse_date_safely(th_raw)

    # Loại bỏ triệt để từ 'vụ án' nếu lọt vào trạng thái tin báo
    if "tạm đình chỉ vụ án" in kq_raw.lower():
        kq_raw = re.sub(r'tạm\s*đình\s*chỉ\s*vụ\s*án', 'Tạm đình chỉ', kq_raw, flags=re.I).strip()
        rec["ket_qua_giai_quyet"] = kq_raw
    elif "đình chỉ vụ án" in kq_raw.lower():
        kq_raw = re.sub(r'đình\s*chỉ\s*vụ\s*án', 'Đình chỉ giải quyết', kq_raw, flags=re.I).strip()
        rec["ket_qua_giai_quyet"] = kq_raw

    is_tam_dc = ("tạm đình chỉ" in kq_raw.lower() or bool(rec.get("ngay_tam_dinh_chi")))

    # 1. Nếu có ngày hết thời hiệu truy cứu TNHS
    if d_th:
        days_th = (d_th - today_date).days
        if days_th < 0:
            rec["so_ngay_con_lai_thoi_hieu"] = f"Đã quá {-days_th} ngày"
            rec["canh_bao_thoi_hieu"] = f"⚠️ Đã quá hạn ({-days_th} ngày)"
            rec["canh_bao_chi_tiet_thoi_hieu"] = "Đã hết thời hiệu truy cứu TNHS"
            rec["so_ngay_con_lai"] = f"Đã quá {-days_th} ngày"
            rec["muc_do_khan_cap"] = f"⚠️ Đã hết thời hiệu truy cứu TNHS (Quá {-days_th} ngày)"
        elif days_th == 0:
            rec["so_ngay_con_lai_thoi_hieu"] = "Hôm nay hết hạn"
            rec["canh_bao_thoi_hieu"] = "🔴 Hết thời hiệu hôm nay (0 ngày)"
            rec["canh_bao_chi_tiet_thoi_hieu"] = "Hết thời hiệu hôm nay"
            rec["so_ngay_con_lai"] = "Hôm nay hết hạn"
            rec["muc_do_khan_cap"] = "🔴 Hết thời hiệu hôm nay (0 ngày)"
        elif days_th <= 7:
            rec["so_ngay_con_lai_thoi_hieu"] = f"Còn {days_th} ngày"
            rec["canh_bao_thoi_hieu"] = f"🔴 Khẩn cấp (Còn {days_th} ngày)"
            rec["canh_bao_chi_tiet_thoi_hieu"] = "Sắp hết thời hiệu truy cứu TNHS"
            rec["so_ngay_con_lai"] = f"Còn {days_th} ngày"
            rec["muc_do_khan_cap"] = f"🔴 Sắp hết thời hiệu truy cứu TNHS (Còn {days_th} ngày)"
        elif days_th <= 15:
            rec["so_ngay_con_lai_thoi_hieu"] = f"Còn {days_th} ngày"
            rec["canh_bao_thoi_hieu"] = f"🟠 Rất gấp (Còn {days_th} ngày)"
            rec["canh_bao_chi_tiet_thoi_hieu"] = "Sắp hết thời hiệu truy cứu TNHS"
            rec["so_ngay_con_lai"] = f"Còn {days_th} ngày"
            rec["muc_do_khan_cap"] = f"🟠 Sắp hết thời hiệu truy cứu TNHS (Còn {days_th} ngày)"
        elif days_th <= 30:
            rec["so_ngay_con_lai_thoi_hieu"] = f"Còn {days_th} ngày"
            rec["canh_bao_thoi_hieu"] = f"🟡 Sắp hết hạn (Còn {days_th} ngày)"
            rec["canh_bao_chi_tiet_thoi_hieu"] = "Sắp hết thời hiệu truy cứu TNHS"
            rec["so_ngay_con_lai"] = f"Còn {days_th} ngày"
            rec["muc_do_khan_cap"] = f"🟡 Sắp hết thời hiệu truy cứu TNHS (Còn {days_th} ngày)"
        else:
            rec["so_ngay_con_lai_thoi_hieu"] = f"Còn {days_th} ngày"
            rec["canh_bao_thoi_hieu"] = f"🟢 Bình thường (Còn {days_th} ngày)"
            rec["canh_bao_chi_tiet_thoi_hieu"] = "Trong thời hiệu truy cứu TNHS"
            if is_tam_dc:
                rec["so_ngay_con_lai"] = f"Còn {days_th} ngày"
                rec["muc_do_khan_cap"] = f"⏸️ Tạm đình chỉ giải quyết (Còn {days_th} ngày TH)"
    elif is_tam_dc:
        rec["so_ngay_con_lai"] = "Tạm đình chỉ"
        rec["muc_do_khan_cap"] = "⏸️ Tạm đình chỉ giải quyết"
        if not rec.get("canh_bao_thoi_hieu"):
            rec["canh_bao_thoi_hieu"] = "⚠️ Tạm đình chỉ (Chưa rõ thời hiệu)"
        if not rec.get("canh_bao_chi_tiet_thoi_hieu"):
            rec["canh_bao_chi_tiet_thoi_hieu"] = "Cần rà soát bổ sung ngày hết thời hiệu"
    else:
        curr_urg = str(rec.get("muc_do_khan_cap") or "")
        if "tạm đình chỉ vụ án" in curr_urg.lower():
            rec["muc_do_khan_cap"] = "⏸️ Tạm đình chỉ giải quyết"
        elif "đình chỉ vụ án" in curr_urg.lower():
            rec["muc_do_khan_cap"] = "⏹️ Đình chỉ giải quyết"


def normalize_case_stage_and_status(rec, filename=""):
    """
    Đảm bảo tính nhất quán và chính xác triệt để 100% giữa:
    1. Giai đoạn tố tụng (giai_doan)
    2. Trạng thái vụ án (trang_thai_vu_an / trang_thai)
    3. Cảnh báo tiến độ / Khẩn cấp (muc_do_khan_cap)
    4. Cảnh báo thời hiệu (canh_bao_thoi_hieu)
    5. Số ngày còn lại (so_ngay_con_lai)

    Quy tắc đồng bộ nghiệp vụ:
    - Nếu có cảnh báo thời hiệu / ngày hết thời hiệu truy cứu TNHS -> Cột cảnh báo tiến độ và số ngày còn lại ưu tiên hiển thị thời hiệu TNHS.
    - Khi trạng thái là 'Tạm đình chỉ' (chưa có thời hiệu) -> Cảnh báo tiến độ: '⏸️ Tạm đình chỉ vụ án'
    - Khi trạng thái là 'Đình chỉ' -> Cảnh báo tiến độ: '⏹️ Đình chỉ vụ án'
    - Khi trạng thái là 'Đã xét xử' -> Cảnh báo tiến độ: '⚖️ Đã xét xử'
    - Khi trạng thái là 'Hoàn thành' / 'Đã giải quyết' / 'Đã kết thúc' -> Cảnh báo tiến độ: '✔ Hoàn thành / Đã giải quyết'
    - Khi trạng thái là 'Phục hồi' -> Cảnh báo tiến độ: '▶️ Phục hồi giải quyết'
    - Khi đang giải quyết bình thường trong hạn -> Cảnh báo tiến độ khớp đúng giai đoạn (Điều tra / Truy tố / Sơ thẩm).
    """
    if not rec:
        return

    # NẾU LÀ TIN BÁO TỘI PHẠM (Tiền khởi tố) -> Chuyển sang chuẩn hóa tin báo (Đ.148 BLTTHS)
    if "ma_tin_bao" in rec or "loai_tin" in rec or "TB_" in str(rec.get("ma_vu_an") or ""):
        normalize_tin_bao_status(rec)
        return
    st_val = str(rec.get("trang_thai_vu_an") or rec.get("trang_thai") or "").strip()
    gd_curr = str(rec.get("giai_doan") or "").strip()
    lt_val = str(rec.get("loai_thu_ly") or "").strip()
    fn_val = filename or str(rec.get("nguon_du_lieu") or "").strip()

    gd_correct = detect_procedural_stage(
        status_text=st_val,
        filename=fn_val,
        loai_thu_ly=lt_val,
        current_stage=gd_curr
    )

    rec["giai_doan"] = gd_correct

    # 1. Chuẩn hóa tiền tố Trạng thái vụ án nếu chưa có
    if st_val and ":" not in st_val and gd_correct in ("Điều tra", "Truy tố", "Sơ thẩm", "Phúc thẩm", "Thi hành án"):
        st_low = st_val.lower()
        if not any(st_low.startswith(x) for x in ("điều tra", "truy tố", "sơ thẩm", "phúc thẩm", "thi hành án")):
            st_val = f"{gd_correct}: {st_val}"
            rec["trang_thai_vu_an"] = st_val

    st_low = st_val.lower()
    lt_low = lt_val.lower()

    has_tam_dinh_chi = (
        "tạm đình chỉ" in st_low 
        or bool(rec.get("ngay_tam_dinh_chi"))
        or bool(rec.get("ly_do_tam_dinh_chi"))
        or "tạm đình chỉ" in lt_low
    )
    has_dinh_chi = (
        "đình chỉ" in st_low and not has_tam_dinh_chi
    )
    has_xet_xu = any(x in st_low for x in ("đã xét xử", "xét xử xong", "bản án", "tuyên án", "chuyển thi hành án"))
    has_hoan_thanh = any(x in st_low for x in (
        "hoàn thành", "đã giải quyết", "kết thúc điều tra", "đã ra cáo trạng", "đã ra bản cáo trạng", 
        "kết luận điều tra", "đã truy tố", "chuyển hồ sơ", "chuyển tòa án",
        "chuyển cơ quan khác", "không khởi tố", "nhập vụ án"
    ))
    has_phuc_hoi = ("phục hồi" in st_low or "phục hồi" in lt_low)

    # 2. ĐỒNG BỘ CẢNH BÁO TIẾN ĐỘ / KHẨN CẤP (muc_do_khan_cap) VÀ SỐ NGÀY CÒN LẠI
    has_statute_warning = update_case_urgency_and_statute(rec, filename=fn_val)

    if not has_statute_warning:
        if has_tam_dinh_chi:
            if "vụ án" in st_low:
                rec["muc_do_khan_cap"] = "⏸️ Tạm đình chỉ vụ án"
            elif "điều tra" in st_low:
                rec["muc_do_khan_cap"] = "⏸️ Tạm đình chỉ điều tra"
            elif "bị can" in st_low:
                rec["muc_do_khan_cap"] = "⏸️ Tạm đình chỉ với bị can"
            else:
                rec["muc_do_khan_cap"] = "⏸️ Tạm đình chỉ vụ án"

            if not rec.get("so_ngay_con_lai") or rec.get("so_ngay_con_lai") == "Đang kiểm sát điều tra":
                rec["so_ngay_con_lai"] = "Tạm đình chỉ"
            if not rec.get("canh_bao_thoi_hieu"):
                rec["canh_bao_thoi_hieu"] = "Tạm đình chỉ (Chưa rõ thời hiệu)"

    elif has_dinh_chi:
        if "vụ án" in st_low:
            rec["muc_do_khan_cap"] = "⏹️ Đình chỉ vụ án"
        elif "điều tra" in st_low:
            rec["muc_do_khan_cap"] = "⏹️ Đình chỉ điều tra"
        elif "bị can" in st_low:
            rec["muc_do_khan_cap"] = "⏹️ Đình chỉ với bị can"
        else:
            rec["muc_do_khan_cap"] = "⏹️ Đình chỉ vụ án"
        rec["so_ngay_con_lai"] = rec.get("so_ngay_con_lai") or "Đã đình chỉ"
        if not rec.get("canh_bao_thoi_hieu"):
            rec["canh_bao_thoi_hieu"] = "Đã đình chỉ giải quyết"

    elif has_xet_xu:
        if "sơ thẩm" in st_low:
            rec["muc_do_khan_cap"] = "⚖️ Đã xét xử sơ thẩm"
        elif "phúc thẩm" in st_low:
            rec["muc_do_khan_cap"] = "⚖️ Đã xét xử phúc thẩm"
        else:
            rec["muc_do_khan_cap"] = "⚖️ Đã xét xử"
        rec["so_ngay_con_lai"] = rec.get("so_ngay_con_lai") or "Đã xét xử"
        if not rec.get("canh_bao_thoi_hieu"):
            rec["canh_bao_thoi_hieu"] = "Đã hoàn thành xét xử"

    elif has_hoan_thanh:
        if "kết thúc điều tra" in st_low:
            rec["muc_do_khan_cap"] = "✔ Đã kết thúc điều tra"
        elif "cáo trạng" in st_low or "đã truy tố" in st_low:
            rec["muc_do_khan_cap"] = "✔ Đã ban hành cáo trạng / Truy tố"
        elif "chuyển tòa" in st_low:
            rec["muc_do_khan_cap"] = "✔ Đã chuyển Tòa án"
        elif "chuyển cơ quan" in st_low:
            rec["muc_do_khan_cap"] = "✔ Đã chuyển cơ quan khác"
        elif "không khởi tố" in st_low:
            rec["muc_do_khan_cap"] = "✔ Không khởi tố vụ án"
        elif "nhập" in st_low:
            rec["muc_do_khan_cap"] = "✔ Đã nhập vụ án"
        else:
            rec["muc_do_khan_cap"] = "✔ Hoàn thành / Đã giải quyết"
        rec["so_ngay_con_lai"] = rec.get("so_ngay_con_lai") or "Hoàn thành"
        if not rec.get("canh_bao_thoi_hieu"):
            rec["canh_bao_thoi_hieu"] = "Đã hoàn thành tố tụng"

    elif has_phuc_hoi:
        if "điều tra" in st_low:
            rec["muc_do_khan_cap"] = "▶️ Phục hồi điều tra"
        else:
            rec["muc_do_khan_cap"] = "▶️ Phục hồi giải quyết"

    else:
        # Trường hợp đang giải quyết bình thường trong hạn
        curr_urg = str(rec.get("muc_do_khan_cap") or "")
        has_symbol_urg = any(sym in curr_urg for sym in ("🔴", "🟠", "🟡", "⚠️", "🟢"))

        if not has_symbol_urg:
            if gd_correct == "Điều tra":
                rec["muc_do_khan_cap"] = "Đang kiểm sát điều tra"
            elif gd_correct == "Truy tố":
                rec["muc_do_khan_cap"] = "Đang THQCT & kiểm sát truy tố"
            elif gd_correct == "Sơ thẩm":
                rec["muc_do_khan_cap"] = "Đang kiểm sát xét xử sơ thẩm"
            elif gd_correct == "Phúc thẩm":
                rec["muc_do_khan_cap"] = "Đang kiểm sát xét xử phúc thẩm"
            elif gd_correct == "Thi hành án":
                rec["muc_do_khan_cap"] = "Đang kiểm sát thi hành án"
            else:
                rec["muc_do_khan_cap"] = f"Đang thụ lý {gd_correct}"

    # 3. Làm sạch triệt để: Ngăn chặn tuyệt đối trạng thái tố tụng / biểu mẫu lọt vào cột Cán bộ / KSV
    for fld in ("can_bo_giai_quyet", "can_bo_thu_ly", "lanh_dao_vien", "lanh_dao_phong"):
        val = str(rec.get(fld) or "").strip()
        if val and (is_procedural_status_text(val) or is_form_status_text(val)):
            curr_tt = str(rec.get("trang_thai_vu_an") or "").strip()
            if not curr_tt or is_form_status_text(curr_tt) or "đang giải quyết" in curr_tt.lower():
                if is_form_status_text(curr_tt) and not rec.get("trang_thai_bieu_mau"):
                    rec["trang_thai_bieu_mau"] = curr_tt
                if is_procedural_status_text(val):
                    rec["trang_thai_vu_an"] = val
            elif is_form_status_text(val) and not rec.get("trang_thai_bieu_mau"):
                rec["trang_thai_bieu_mau"] = val
            rec[fld] = ""

    # 4. Ngăn chặn tuyệt đối tình trạng cột Trạng thái vụ án bị gán nhầm tên cán bộ,
    # trong khi cột Biểu mẫu lại đang giữ Trạng thái tố tụng thực sự do dữ liệu bị lệch cột:
    tt_val = str(rec.get("trang_thai_vu_an") or "").strip()
    bm_val = str(rec.get("trang_thai_bieu_mau") or "").strip()
    if is_procedural_status_text(bm_val) and not is_procedural_status_text(tt_val):
        if tt_val and not is_form_status_text(tt_val):
            ld_v, ld_p, ksv, cb_tl = split_officers(tt_val, "")
            if ld_v and not rec.get("lanh_dao_vien"):
                rec["lanh_dao_vien"] = ld_v
            if ld_p and not rec.get("lanh_dao_phong"):
                rec["lanh_dao_phong"] = ld_p
            if ksv and not rec.get("can_bo_giai_quyet"):
                rec["can_bo_giai_quyet"] = ksv
            if cb_tl and not rec.get("can_bo_thu_ly"):
                rec["can_bo_thu_ly"] = cb_tl
        rec["trang_thai_vu_an"] = bm_val
        rec["trang_thai_bieu_mau"] = ""

    return rec


def parse_case_header_to_col_map(row):
    """
    Phân tích linh hoạt tiêu đề cột thành map vị trí tương ứng cho Sổ/Báo cáo vụ án QLAHS.
    Hỗ trợ mọi biến thể tiêu đề của QLAHS từ trực tuyến, kết xuất danh sách, hoặc copy tay.
    """
    r_strs = [str(c or "").strip().lower() for c in row]
    col_map = {
        "va": 1,
        "tl": None,
        "bc": None,
        "td": None,
        "lt": None,
        "cb": None,
        "tt": None,
        "bm": None,
        "dv_moi": None,
    }
    for idx, c_text in enumerate(r_strs):
        if not c_text:
            continue
        if "thông tin vụ án" in c_text or "mã vụ án" in c_text:
            col_map["va"] = idx
        elif "thông tin thụ lý" in c_text:
            col_map["tl"] = idx
        elif "loại thụ lý" in c_text:
            if col_map["tl"] is None:
                col_map["tl"] = idx
        elif "thông tin bị can" in c_text or "bị can" in c_text:
            col_map["bc"] = idx
        elif "loại tội" in c_text:
            col_map["lt"] = idx
        elif ("tội danh" in c_text or c_text == "tội" or "thông tin tội danh" in c_text) and "loại tội" not in c_text:
            col_map["td"] = idx
        elif "cán bộ" in c_text or "kiểm sát viên" in c_text:
            col_map["cb"] = idx
        elif "trạng thái vụ án" in c_text:
            col_map["tt"] = idx
        elif "trạng thái biểu mẫu" in c_text or "biểu mẫu" in c_text:
            col_map["bm"] = idx
        elif "mới nhất" in c_text or "cập nhật mới" in c_text or "đơn vị cập nhật" in c_text:
            col_map["dv_moi"] = idx
    return col_map


def detect_initial_case_col_map(raw_rows):
    """Xác định cấu trúc cột ban đầu từ hàng tiêu đề hoặc phân tích nội dung dữ liệu 35 dòng đầu."""
    for r in raw_rows[:25]:
        if not r or not isinstance(r, (list, tuple)):
            continue
        line_str = " ".join(str(c or "").strip().lower() for c in r)
        if "thông tin vụ án" in line_str or "mã vụ án" in line_str:
            return parse_case_header_to_col_map(r)

    col_map = {
        "va": 1,
        "tl": None,
        "bc": None,
        "td": None,
        "lt": None,
        "cb": None,
        "tt": None,
        "bm": None,
    }
    col_scores = {}
    for r in raw_rows[:35]:
        if not r or not isinstance(r, (list, tuple)):
            continue
        for idx, val in enumerate(r):
            if val is None:
                continue
            v_s = str(val).strip().replace("\xa0", " ")
            v_low = v_s.lower()
            if not v_s:
                continue
            if idx not in col_scores:
                col_scores[idx] = set()
            if v_s.startswith("VA_") or "số qđ khởi tố:" in v_low:
                col_scores[idx].add("va")
            if "số thụ lý:" in v_low or "ngày thụ lý:" in v_low or "thụ lý mới" in v_low:
                col_scores[idx].add("tl")
            if any(x in v_low for x in ("đặc biệt nghiêm trọng", "ít nghiêm trọng", "rất nghiêm trọng", "nghiêm trọng")):
                col_scores[idx].add("lt")
            if idx != 1 and ("điều " in v_low or "khoản " in v_low or "tội " in v_low):
                col_scores[idx].add("td")
            if "cán bộ thụ lý:" in v_low or "cán bộ giải quyết:" in v_low:
                col_scores[idx].add("cb")
            if is_procedural_status_text(v_s):
                col_scores[idx].add("tt")
            if is_form_status_text(v_s):
                col_scores[idx].add("bm")

    for idx, tags in col_scores.items():
        if "va" in tags and col_map["va"] is None:
            col_map["va"] = idx
        if "tl" in tags and col_map["tl"] is None:
            col_map["tl"] = idx
        if "lt" in tags and col_map["lt"] is None:
            col_map["lt"] = idx
        if "td" in tags and col_map["td"] is None:
            col_map["td"] = idx
        if "cb" in tags and col_map["cb"] is None:
            col_map["cb"] = idx
        if "tt" in tags and col_map["tt"] is None:
            col_map["tt"] = idx
        if "bm" in tags and col_map["bm"] is None:
            col_map["bm"] = idx

    return col_map


def parse_qlahs_case_report_rows(raw_rows, progress_cb=None, filename=""):
    file_context_unit = detect_file_unit(filename, raw_rows)
    """Bóc tách Báo cáo danh sách hồ sơ vụ án (8 cột: STT, Thông tin vụ án, Loại tội, Tội, Thông tin bị can, Loại thụ lý, Trạng thái vụ án, Trạng thái biểu mẫu)."""
    cases = []
    current_col_map = detect_initial_case_col_map(raw_rows)

    for r_i, row in enumerate(raw_rows, 1):
        if not row or not any(row):
            continue
        line_str = " ".join(str(c or "").lower() for c in row)
        if "thông tin vụ án" in line_str or "mã vụ án" in line_str:
            current_col_map = parse_case_header_to_col_map(row)
            continue
        c0 = str(row[0] if len(row) > 0 and row[0] is not None else "").strip()
        if c0.lower() == "stt" or "thời gian kết xuất" in line_str or "danh sách hồ sơ" in line_str:
            continue

        def get_c(c_idx):
            if c_idx is not None and len(row) > c_idx and row[c_idx] is not None:
                return str(row[c_idx]).strip().replace("\xa0", " ")
            return ""

        c1 = get_c(current_col_map.get("va"))
        if not c1 or c1.lower() == "thông tin vụ án":
            continue

        c2 = get_c(current_col_map.get("lt"))
        c3 = get_c(current_col_map.get("td"))
        c4 = get_c(current_col_map.get("bc"))
        c5 = get_c(current_col_map.get("tl"))
        c_cb = get_c(current_col_map.get("cb"))
        c6 = get_c(current_col_map.get("tt"))
        c7 = get_c(current_col_map.get("bm"))
        c_dv_moi = get_c(current_col_map.get("dv_moi"))

        # Lưới an toàn tự chữa lành nếu dữ liệu bị lệch cột
        if is_procedural_status_text(c_cb):
            if is_form_status_text(c6):
                c7 = c6
            c6 = c_cb
            c_cb = ""
        if is_procedural_status_text(c7) and not is_procedural_status_text(c6) and c6:
            c_cb = c6
            c6 = c7
            c7 = ""
        if is_form_status_text(c6) and not c7:
            c7 = c6
            c6 = ""

        lines = [ln.strip() for ln in c1.replace("\r\n", "\n").split("\n") if ln.strip()]
        ma_va = lines[0] if lines else ""
        ten_va = []
        so_qd = ""
        ngay_qd = ""
        ngay_tn = ""
        thoi_han = ""
        ngay_het_th = ""
        canh_bao = ""

        for ln in lines[1:]:
            ln_low = ln.lower()
            if ln_low.startswith("mã số qđ khởi tố vụ án:") or ln_low.startswith("số qđ khởi tố:"):
                so_qd = ln.split(":", 1)[1].strip()
            elif ln_low.startswith("ngày ra qđ khởi tố:"):
                ngay_qd = ln.split(":", 1)[1].strip()
            elif ln_low.startswith("ngày tiếp nhận hồ sơ:"):
                ngay_tn = ln.split(":", 1)[1].strip()
            elif ln_low.startswith("thời hạn xử lý:") or ln_low.startswith("thời hạn giải quyết:"):
                thoi_han = ln.split(":", 1)[1].strip()
            elif ln_low.startswith("ngày hết thời hiệu"):
                ngay_het_th = ln.split(":", 1)[1].strip()
            elif "quá hạn" in ln_low or "thời hiệu" in ln_low or "sắp hết" in ln_low:
                canh_bao = ln
            else:
                ten_va.append(ln)

        full_ten_va = " - ".join(ten_va)
        bc_dau_vu = ""
        tong_bc = ""
        if c4:
            for bc_ln in c4.split("\n"):
                bc_low = bc_ln.lower()
                if "bị can đầu vụ:" in bc_low:
                    bc_dau_vu = bc_ln.split(":", 1)[1].strip()
                elif "tổng số bị can:" in bc_low:
                    tong_bc = bc_ln.split(":", 1)[1].strip()
                elif not bc_dau_vu:
                    bc_dau_vu = bc_ln.strip()

        prov, vks_tl = parse_case_code_unit(ma_va, file_context_unit)
        if c_dv_moi:
            u_clean = clean_vks_unit_name(c_dv_moi)
            if u_clean and "chưa" not in u_clean.lower():
                vks_tl = u_clean
        else:
            vks_tl = clean_vks_unit_name(vks_tl)
        ksv_from_text = extract_officers_from_text(full_ten_va)
        raw_officer_src = c_cb if c_cb else ksv_from_text
        ld_v, ld_p, ksv, cb_tl = split_officers(raw_officer_src, "")

        # Nhận diện chính xác tuyệt đối Giai đoạn tố tụng từ trạng thái vụ án, tên file và loại thụ lý
        tt_va = c6
        gd_det = detect_procedural_stage(status_text=tt_va, filename=filename, loai_thu_ly=c5)
        if not tt_va:
            tt_va = f"{gd_det}: Đang giải quyết" if gd_det != "Điều tra" else "Điều tra: Đang THQCT và kiểm sát điều tra"

        tt_va_low = tt_va.lower()
        if "tạm đình chỉ" in tt_va_low:
            default_muc_do = "⏸️ Tạm đình chỉ vụ án"
        elif "đình chỉ" in tt_va_low:
            default_muc_do = "⏹️ Đình chỉ vụ án"
        elif any(x in tt_va_low for x in ("đã xét xử", "xét xử xong")):
            default_muc_do = "⚖️ Đã xét xử"
        elif any(x in tt_va_low for x in ("hoàn thành", "đã giải quyết", "kết thúc điều tra", "cáo trạng", "đã truy tố")):
            default_muc_do = "✔ Hoàn thành / Đã giải quyết"
        elif gd_det == "Truy tố":
            default_muc_do = "Đang THQCT & kiểm sát truy tố"
        elif gd_det == "Sơ thẩm":
            default_muc_do = "Đang kiểm sát xét xử sơ thẩm"
        elif gd_det == "Phúc thẩm":
            default_muc_do = "Đang kiểm sát xét xử phúc thẩm"
        elif gd_det == "Thi hành án":
            default_muc_do = "Đang kiểm sát thi hành án"
        else:
            default_muc_do = "Đang kiểm sát điều tra"

        muc_do = f"⚠️ {canh_bao}" if canh_bao else default_muc_do

        rec_item = {
            "stt": len(cases) + 1,
            "nguon_du_lieu": filename or "DS1: Hồ sơ vụ án",
            "ma_vu_an": ma_va,
            "ten_vu_an": full_ten_va,
            "tinh_thanh": prov,
            "so_qd_khoi_to": so_qd,
            "ngay_qd_khoi_to": ngay_qd,
            "don_vi_khoi_to": "Công an Thành phố Hồ Chí Minh",
            "so_thu_ly": "",
            "ngay_thu_ly": ngay_tn,
            "loai_thu_ly": c5 or "Thụ lý mới",
            "vks_thu_ly": vks_tl,
            "ngay_tiep_nhan": ngay_tn,
            "thoi_han_giai_quyet": thoi_han,
            "ngay_het_thoi_hieu": ngay_het_th,
            "so_ngay_con_lai": "",
            "muc_do_khan_cap": muc_do,
            "canh_bao_thoi_hieu": canh_bao or "",
            "bi_can_dau_vu": bc_dau_vu,
            "tong_bi_can": tong_bc or ("1" if bc_dau_vu else "0"),
            "dieu_luat": c3,
            "loai_toi": c2,
            "lanh_dao_vien": ld_v,
            "lanh_dao_phong": ld_p,
            "can_bo_giai_quyet": ksv,
            "can_bo_thu_ly": cb_tl,
            "giai_doan": gd_det,
            "trang_thai_vu_an": tt_va,
            "trang_thai_bieu_mau": c7,
            "ten_don_vi_moi_nhat": c_dv_moi or ""
        }
        normalize_case_stage_and_status(rec_item, filename=filename)
        cases.append(rec_item)

    return cases

def reconcile_and_merge_multiple_case_lists(parsed_files, progress_cb=None):
    """
    Rà soát, đối chiếu và hợp nhất nhiều file QLAHS (2, 3, 4, ... N files).
    Tự động so khớp theo Mã vụ án & Số QĐ khởi tố, bù đắp thông tin còn thiếu
    từ tất cả các file để tạo danh sách hợp nhất 27 cột đầy đủ 100%.
    
    parsed_files: danh sách các tuple (tên_file, định_dạng, danh_sách_bản_ghi)
    """
    def norm_key(s):
        return re.sub(r'[^a-zA-Z0-9]', '', str(s or '')).upper()

    master_records = []
    index_by_code = {}  # code_key -> index in master_records
    index_by_qd = {}    # qd_key -> index in master_records
    file_stats = []     # (fname, total, matched, new)
    total_matches = 0

    for file_idx, (fname, fmt, records) in enumerate(parsed_files):
        if progress_cb:
            progress_cb(
                f"Đang so khớp & bù đắp dữ liệu Tập tin {file_idx + 1}/{len(parsed_files)}...",
                f"{fname} ({len(records):,} vụ án)"
            )

        file_matched = 0
        file_new = 0

        for r in records:
            code_k = norm_key(r.get("ma_vu_an"))
            qd_raw = str(r.get("so_qd_khoi_to") or "").strip()
            dv_raw = str(r.get("don_vi_khoi_to") or r.get("vks_thu_ly") or "").strip()
            ngay_qd = str(r.get("ngay_qd_khoi_to") or "")
            m_yr = re.search(r'\b(20\d\d)\b', ngay_qd)
            nam_qd = m_yr.group(1) if m_yr else ""

            # Khóa phụ chỉ dùng khi số QĐ có độ dài hoặc đi kèm đơn vị & năm (tránh trùng số thứ tự nhỏ như 1, 2, 31, 101)
            qd_k = ""
            if qd_raw and len(qd_raw) >= 2:
                if nam_qd or dv_raw:
                    qd_k = norm_key(qd_raw + nam_qd + dv_raw)
                elif len(qd_raw) >= 5:  # Ví dụ C1121/QĐ-CSHS
                    qd_k = norm_key(qd_raw)

            target_idx = None
            # ƯU TIÊN SỐ 1 TUYỆT ĐỐI: Khớp theo Mã vụ án chính thức
            if code_k:
                if code_k in index_by_code:
                    target_idx = index_by_code[code_k]
                # Nếu có mã vụ án nhưng chưa có trong danh sách -> CHẮC CHẮN LÀ VỤ ÁN MỚI
                # Tuyệt đối không được gộp vào vụ khác chỉ vì trùng số QĐ khởi tố!
            elif qd_k and qd_k in index_by_qd:
                # Chỉ xét khớp số QĐ khởi tố khi bản ghi này KHÔNG CÓ mã vụ án
                target_idx = index_by_qd[qd_k]

            if target_idx is not None:
                # Đã có trong danh sách -> Tiến hành bù đắp dữ liệu
                target = master_records[target_idx]
                file_matched += 1
                total_matches += 1

                sources = target.get("_sources", [])
                if fname not in sources:
                    sources.append(fname)
                    target["_sources"] = sources
                if len(sources) > 1:
                    target["nguon_du_lieu"] = f"Khớp nối ({len(sources)} file): " + " + ".join(sources)
                else:
                    target["nguon_du_lieu"] = sources[0]

                # Bù đắp & làm giàu dữ liệu thông minh giữa các file
                PLACEHOLDER_VALS = {
                    "chưa xác định đơn vị", "chưa xác định đơn vị (mã cũ qlahs)",
                    "khác / chưa rõ", "chưa xác định", "thụ lý mới",
                    "đang kiểm sát điều tra", "none", "null", "",
                    "tạm đình chỉ", "tạm đình chỉ (chưa rõ thời hiệu)", "chưa rõ thời hiệu",
                    "chưa rõ", "chưa phân loại",
                    "⏸️ tạm đình chỉ vụ án", "⏸️ tạm đình chỉ điều tra", "⏸️ tạm đình chỉ với bị can"
                }

                for k, v in r.items():
                    if k.startswith("_") or k in ("stt", "nguon_du_lieu"):
                        continue
                    val_new = str(v or "").strip()
                    if not val_new:
                        continue
                    val_curr = str(target.get(k) or "").strip()

                    # Trường hiện tại rỗng hoặc chỉ là giá trị giữ chỗ -> Cập nhật ngay
                    if not val_curr or val_curr.lower() in PLACEHOLDER_VALS:
                        target[k] = v
                        continue

                    # Ngày hết thời hiệu: Luôn ưu tiên giá trị ngày hợp lệ nếu có
                    if k == "ngay_het_thoi_hieu":
                        if parse_date_safely(val_new) and not parse_date_safely(val_curr):
                            target[k] = val_new
                        continue

                    # Tên đơn vị cập nhật mới nhất: Ưu tiên TUYỆT ĐỐI số 1
                    elif k in ("ten_don_vi_moi_nhat", "don_vi_moi_nhat"):
                        if val_new and val_new.lower() not in PLACEHOLDER_VALS:
                            target["ten_don_vi_moi_nhat"] = val_new
                            u_std = clean_vks_unit_name(val_new)
                            if u_std and "chưa" not in u_std.lower():
                                target["vks_thu_ly"] = u_std
                        continue

                    # Cảnh báo thời hiệu / quá hạn: Luôn ưu tiên cảnh báo thực tế hơn giá trị mặc định
                    elif k in ("canh_bao_thoi_hieu", "canh_bao_qua_han"):
                        if any(x in val_new.lower() for x in ("sắp hết", "hết thời hiệu", "quá hạn")):
                            target[k] = val_new
                        continue

                    # Đơn vị VKS thụ lý: Ưu tiên đơn vị cụ thể (Khu vực X hoặc Phòng Y) hơn đơn vị chung chung
                    elif k == "vks_thu_ly":
                        # Nếu đã có đơn vị cập nhật mới nhất thì bảo lưu tuyệt đối
                        if target.get("ten_don_vi_moi_nhat"):
                            continue
                        is_new_specific = ("khu vực" in val_new.lower() or "phòng" in val_new.lower() or "p2" in val_new.lower())
                        is_curr_specific = ("khu vực" in val_curr.lower() or "phòng" in val_curr.lower() or "p2" in val_curr.lower())
                        if is_new_specific and not is_curr_specific:
                            target[k] = clean_vks_unit_name(v)
                        elif is_new_specific and is_curr_specific:
                            target[k] = clean_vks_unit_name(v)
                        else:
                            target[k] = clean_vks_unit_name(target.get(k))

                    # Tỉnh / Thành phố: Ưu tiên TP. Hồ Chí Minh nếu có
                    elif k == "tinh_thanh":
                        if "hồ chí minh" in val_new.lower() and "hồ chí minh" not in val_curr.lower():
                            target[k] = v

                    # Các trường mô tả chi tiết: Ưu tiên chuỗi dài hơn, chi tiết hơn
                    elif k in ("ten_vu_an", "dieu_luat", "bi_can_dau_vu", "don_vi_khoi_to"):
                        if len(val_new) > len(val_curr):
                            target[k] = v

                    # Phân loại tội: Ưu tiên thông tin rõ ràng
                    elif k == "loai_toi":
                        if val_curr.lower() in ("chưa rõ", "khác", "") and val_new.lower() not in ("chưa rõ", "khác", ""):
                            target[k] = v

                # Hợp nhất thông minh Giai đoạn tố tụng & Trạng thái vụ án theo tiến trình tố tụng hình sự
                STAGE_PRECEDENCE = {
                    "Tin báo": 1,
                    "Điều tra": 2,
                    "Truy tố": 3,
                    "Sơ thẩm": 4,
                    "Phúc thẩm": 5,
                    "Thi hành án": 6,
                    "GĐT/TT": 7,
                }
                new_gd = r.get("giai_doan") or detect_procedural_stage(r.get("trang_thai_vu_an"), fname, r.get("loai_thu_ly"))
                curr_gd = target.get("giai_doan") or detect_procedural_stage(target.get("trang_thai_vu_an"), target.get("nguon_du_lieu"), target.get("loai_thu_ly"))

                new_p = STAGE_PRECEDENCE.get(new_gd, 0)
                curr_p = STAGE_PRECEDENCE.get(curr_gd, 0)

                # Nếu file mới mang giai đoạn tố tụng tiến triển hơn (hoặc bản ghi hiện tại chưa rõ) -> Cập nhật giai đoạn & trạng thái mới nhất
                if new_p > curr_p or not curr_gd or curr_gd.lower() in ("chưa rõ", "khác", ""):
                    target["giai_doan"] = new_gd
                    if r.get("trang_thai_vu_an"):
                        target["trang_thai_vu_an"] = r["trang_thai_vu_an"]
                elif new_p == curr_p:
                    # Cùng giai đoạn: ưu tiên trạng thái chi tiết hơn hoặc có tiền tố chuẩn
                    new_st = str(r.get("trang_thai_vu_an") or "")
                    curr_st = str(target.get("trang_thai_vu_an") or "")
                    if len(new_st) > len(curr_st) or (":" in new_st and ":" not in curr_st):
                        target["trang_thai_vu_an"] = new_st

                normalize_case_stage_and_status(target, fname)

                for fld in ("lanh_dao_vien", "lanh_dao_phong", "can_bo_giai_quyet", "can_bo_thu_ly"):
                    val_fld_new = r.get(fld)
                    if val_fld_new:
                        val_fld_curr = target.get(fld)
                        if not val_fld_curr:
                            target[fld] = val_fld_new
                        else:
                            names_curr = [clean_person_name(x) for x in str(val_fld_curr).split(";") if clean_person_name(x)]
                            names_new = [clean_person_name(x) for x in str(val_fld_new).split(";") if clean_person_name(x)]
                            merged_names = list(dict.fromkeys(names_curr + names_new))
                            target[fld] = "; ".join(merged_names)

                # Ưu tiên ghi nhận cảnh báo nếu file mới có
                if r.get("canh_bao_qua_han"):
                    target["canh_bao_thoi_hieu"] = r["canh_bao_qua_han"]
                    target["muc_do_khan_cap"] = f"⚠️ {r['canh_bao_qua_han']}"
                elif r.get("canh_bao_thoi_hieu") and not target.get("canh_bao_thoi_hieu"):
                    target["canh_bao_thoi_hieu"] = r["canh_bao_thoi_hieu"]
                    target["muc_do_khan_cap"] = r.get("muc_do_khan_cap") or f"⚠️ {r['canh_bao_thoi_hieu']}"

                if code_k and code_k not in index_by_code:
                    index_by_code[code_k] = target_idx
                if qd_k and qd_k not in index_by_qd:
                    index_by_qd[qd_k] = target_idx
            else:
                # Vụ án mới -> Thêm vào master
                file_new += 1
                new_rec = dict(r)
                new_rec["_sources"] = [fname]
                new_rec["nguon_du_lieu"] = fname
                if new_rec.get("canh_bao_qua_han"):
                    new_rec["canh_bao_thoi_hieu"] = new_rec["canh_bao_qua_han"]
                    new_rec["muc_do_khan_cap"] = f"⚠️ {new_rec['canh_bao_qua_han']}"

                normalize_case_stage_and_status(new_rec, fname)
                new_idx = len(master_records)
                master_records.append(new_rec)

                if code_k:
                    index_by_code[code_k] = new_idx
                if qd_k:
                    index_by_qd[qd_k] = new_idx

        file_stats.append({
            "fname": fname,
            "total": len(records),
            "matched": file_matched,
            "new": file_new,
        })

    # Cập nhật số thứ tự STT chuẩn hóa, dọn dẹp biến nội bộ và gán đơn vị VKS thụ lý
    for idx, r in enumerate(master_records, 1):
        r["stt"] = idx
        r.pop("_sources", None)

        src = str(r.get("nguon_du_lieu") or "")
        unit_from_src = extract_unit_from_source(src)

        code = str(r.get("ma_vu_an") or "")
        is_new_code = bool(
            re.search(r'VA_E01\.59[._]\d{3}_\d{4}_\d+', code, re.I) or 
            re.search(r'VA_E01\.59_\d{4}_\d+', code, re.I)
        )

        curr_u = str(r.get("vks_thu_ly") or "").strip()
        is_unspecific = (
            not curr_u or 
            "chưa xác định" in curr_u.lower() or 
            "chưa rõ" in curr_u.lower() or 
            "mã cũ" in curr_u.lower() or
            curr_u.lower() in ("tp. hồ chí minh", "vksnd thành phố hồ chí minh", "vksnd tp.hcm")
        )

        # ƯU TIÊN SỐ 1 TUYỆT ĐỐI: Kế thừa đơn vị cập nhật mới nhất do người dùng rà soát bổ sung
        if r.get("ten_don_vi_moi_nhat") and str(r.get("ten_don_vi_moi_nhat")).strip():
            u_latest = clean_vks_unit_name(r["ten_don_vi_moi_nhat"])
            if u_latest and "chưa" not in u_latest.lower():
                r["vks_thu_ly"] = u_latest
        elif unit_from_src and (not is_new_code or is_unspecific):
            r["vks_thu_ly"] = unit_from_src
        elif r.get("vks_thu_ly"):
            r["vks_thu_ly"] = clean_vks_unit_name(r["vks_thu_ly"])
        elif unit_from_src:
            r["vks_thu_ly"] = unit_from_src
        else:
            r["vks_thu_ly"] = "Chưa rõ đơn vị VKS"

        # Chuẩn hóa triệt để giai đoạn tố tụng đồng bộ với trạng thái vụ án
        normalize_case_stage_and_status(r)

    return master_records, total_matches, file_stats


def reconcile_and_merge_case_lists(list_a, list_b):
    """Tương thích ngược đối chiếu 2 file"""
    merged, match_n, _ = reconcile_and_merge_multiple_case_lists([
        ("File 1", "CASE", list_a),
        ("File 2", "CASE", list_b)
    ])
    return merged, match_n


def reconcile_and_merge_multiple_tin_bao_lists(parsed_files, progress_cb=None):
    """
    Rà soát, đối chiếu và hợp nhất nhiều file Tin báo tội phạm QLAHS (2, 3, ... N files).
    Tự động so khớp theo Mã tin báo & Số thụ lý (Số TL) + Đơn vị VKS, bù đắp thông tin còn thiếu
    từ tất cả các file để tạo danh sách hợp nhất 36 cột đầy đủ nhất.
    """
    def norm_key(s):
        return re.sub(r'[^a-zA-Z0-9]', '', str(s or '')).upper()

    master_records = []
    index_by_code = {}   # norm_key(ma_tin_bao) -> idx in master_records
    index_by_stl = {}    # norm_key(so_thu_ly + vks_thu_ly) -> idx in master_records
    file_stats = []      # {"fname": fname, "total": ..., "matched": ..., "new": ...}
    total_matches = 0

    for file_idx, (fname, fmt, records) in enumerate(parsed_files):
        if progress_cb:
            progress_cb(
                f"Đang so khớp & bù đắp dữ liệu Tin báo Tập tin {file_idx + 1}/{len(parsed_files)}...",
                f"{fname} ({len(records):,} tin báo)"
            )

        file_matched = 0
        file_new = 0

        for r in records:
            code_k = norm_key(r.get("ma_tin_bao"))
            so_tl = str(r.get("so_thu_ly", "") or "").strip()
            vks_tl = str(r.get("vks_thu_ly", "") or "").strip()
            stl_k = norm_key(so_tl + vks_tl) if so_tl else ""

            target_idx = None
            # ƯU TIÊN SỐ 1 TUYỆT ĐỐI: Khớp theo Mã tin báo
            if code_k:
                if code_k in index_by_code:
                    target_idx = index_by_code[code_k]
                # Có mã tin báo riêng biệt -> Tuyệt đối không gộp nhầm
            elif stl_k and stl_k in index_by_stl:
                # Chỉ xét số thụ lý + đơn vị khi bản ghi này KHÔNG CÓ mã tin báo
                target_idx = index_by_stl[stl_k]

            if target_idx is not None:
                # Đã có trong danh sách -> Tiến hành bù đắp & làm giàu dữ liệu
                target = master_records[target_idx]
                file_matched += 1
                total_matches += 1

                sources = target.get("_sources", [])
                if fname not in sources:
                    sources.append(fname)
                    target["_sources"] = sources
                if len(sources) > 1:
                    target["nguon_du_lieu"] = f"Khớp nối ({len(sources)} file): " + " + ".join(sources)
                else:
                    target["nguon_du_lieu"] = f"Tệp: {sources[0]}"

                # Bù đắp & làm giàu dữ liệu tin báo thông minh
                PLACEHOLDER_VALS = {
                    "chưa xác định đơn vị", "chưa xác định đơn vị (mã cũ qlahs)",
                    "khác / chưa rõ", "chưa xác định", "none", "null", ""
                }

                for k, v in r.items():
                    if k.startswith("_") or k in ("stt", "nguon_du_lieu"):
                        continue
                    val_new = str(v or "").strip()
                    if not val_new:
                        continue
                    val_curr = str(target.get(k) or "").strip()

                    if not val_curr or val_curr.lower() in PLACEHOLDER_VALS:
                        target[k] = v
                        continue

                    if k == "ten_don_vi_moi_nhat":
                        target["vks_thu_ly"] = clean_vks_unit_name(val_new)
                    elif k == "vks_thu_ly":
                        is_new_specific = ("khu vực" in val_new.lower() or "phòng" in val_new.lower() or "p2" in val_new.lower())
                        is_curr_specific = ("khu vực" in val_curr.lower() or "phòng" in val_curr.lower() or "p2" in val_curr.lower())
                        if is_new_specific and not is_curr_specific:
                            target[k] = clean_vks_unit_name(v)
                        elif is_new_specific and is_curr_specific:
                            target[k] = clean_vks_unit_name(v)
                        else:
                            target[k] = clean_vks_unit_name(target.get(k))
                    elif k == "noi_dung_tin_bao" and len(val_new) > len(val_curr):
                        target["noi_dung_tin_bao"] = v
                    elif k in ("ksv_chinh", "ksv_phoi_hop", "dieu_tra_vien", "ld_vien_phu_trach", "ld_phong_phu_trach"):
                        if val_new and (not val_curr or val_curr.startswith("=")):
                            target[k] = v
                        elif val_new and val_curr:
                            names_curr = [clean_person_name(x) for x in val_curr.split(";") if clean_person_name(x)]
                            names_new = [clean_person_name(x) for x in val_new.split(";") if clean_person_name(x)]
                            target[k] = "; ".join(dict.fromkeys(names_curr + names_new))
                    elif k in ("thoi_han_den_ngay", "muc_do_khan_cap", "ngay_tam_dinh_chi", "ly_do_tam_dinh_chi",
                               "ngay_het_thoi_hieu", "so_ngay_con_lai_thoi_hieu", "canh_bao_thoi_hieu", "canh_bao_chi_tiet_thoi_hieu") and val_new and not val_curr:
                        target[k] = v
                    elif k == "canh_bao_thoi_hieu" and "🔴" in val_new and "🔴" not in val_curr:
                        target[k] = v
                        if r.get("canh_bao_chi_tiet_thoi_hieu"):
                            target["canh_bao_chi_tiet_thoi_hieu"] = r["canh_bao_chi_tiet_thoi_hieu"]
                        if r.get("so_ngay_con_lai_thoi_hieu"):
                            target["so_ngay_con_lai_thoi_hieu"] = r["so_ngay_con_lai_thoi_hieu"]

                if code_k and code_k not in index_by_code:
                    index_by_code[code_k] = target_idx
                if stl_k and stl_k not in index_by_stl:
                    index_by_stl[stl_k] = target_idx
            else:
                # Tin báo mới -> Thêm vào master
                file_new += 1
                new_rec = dict(r)
                new_rec["_sources"] = [fname]
                new_rec["nguon_du_lieu"] = f"Tệp: {fname}"

                new_idx = len(master_records)
                master_records.append(new_rec)

                if code_k:
                    index_by_code[code_k] = new_idx
                if stl_k:
                    index_by_stl[stl_k] = new_idx

        file_stats.append({
            "fname": fname,
            "total": len(records),
            "matched": file_matched,
            "new": file_new,
        })

    # Cập nhật STT, dọn biến tạm và gán đơn vị VKS thụ lý từ nguồn đối chiếu
    for idx, r in enumerate(master_records, 1):
        r["stt"] = idx
        r.pop("_sources", None)

        # Kế thừa đơn vị nếu có nguồn cập nhật
        dv_moi = str(r.get("ten_don_vi_moi_nhat") or "").strip()
        if dv_moi:
            u_std = clean_vks_unit_name(dv_moi)
            if u_std and "chưa" not in u_std.lower():
                r["vks_thu_ly"] = u_std
            r.pop("ten_don_vi_moi_nhat", None)
            continue
        r.pop("ten_don_vi_moi_nhat", None)

        src = str(r.get("nguon_du_lieu") or "")
        unit_from_src = extract_unit_from_source(src)

        code = str(r.get("ma_tin_bao") or "")
        is_new_code = bool(
            re.search(r'TB_E01\.59[._]\d{3}_\d{4}_\d+', code, re.I) or 
            re.search(r'TB_E01\.59_\d{4}_\d+', code, re.I)
        )

        curr_u = str(r.get("vks_thu_ly") or "").strip()
        is_unspecific = (
            not curr_u or 
            "chưa xác định" in curr_u.lower() or 
            "chưa rõ" in curr_u.lower() or 
            "mã cũ" in curr_u.lower() or
            curr_u.lower() in ("tp. hồ chí minh", "vksnd thành phố hồ chí minh", "vksnd tp.hcm")
        )

        if unit_from_src and (not is_new_code or is_unspecific):
            r["vks_thu_ly"] = unit_from_src
        elif r.get("vks_thu_ly") and "chưa" not in str(r.get("vks_thu_ly")).lower():
            r["vks_thu_ly"] = clean_vks_unit_name(r["vks_thu_ly"])
        elif unit_from_src:
            r["vks_thu_ly"] = unit_from_src
        else:
            u_code = parse_tin_bao_code_unit(code, unit_from_src)
            r["vks_thu_ly"] = u_code or "Chưa xác định đơn vị (Mã cũ QLAHS)"

        # Chuẩn hóa tiến độ & cảnh báo thời hiệu tin báo tội phạm (Đ.148 BLTTHS)
        normalize_tin_bao_status(r)

    return master_records, total_matches, file_stats


def reconcile_and_link_cases_defendants(case_files, defendant_files, progress_cb=None):
    """
    Rà soát, đối chiếu và LIÊN THÔNG DỮ LIỆU SỔ VỤ ÁN ⟷ DANH SÁCH BỊ CAN QLAHS.
    - Hợp nhất dữ liệu các tệp Vụ án (nếu có nhiều hơn 1 file).
    - Gom nhóm các bị can thành từng cụm/khối đồng phạm vụ án (theo header block hoặc mã vụ án).
    - Tự động liên thông 2 chiều:
      1. Khớp theo Mã vụ án mới (VD: VA_E01.59_2606_1297_001 -> VA_E01.59_2606_1297).
      2. Khớp theo Số CCCD/CMND bị can hoặc tên Bị can đầu vụ trong hồ sơ vụ án.
      3. Khớp đồng phạm: Khi 1 bị can trong khối khớp vụ án, toàn bộ đồng phạm trong khối cùng thuộc vụ án đó.
    - Làm giàu Sổ vụ án:
      * Bổ sung 'tong_bi_can' chuẩn xác.
      * Bổ sung 'bi_can_dau_vu' kèm CCCD.
      * Bổ sung cột 'danh_sach_bi_can_chi_tiet' đánh số thứ tự từng bị can trong vụ.
    - Phân tách bị can:
      * linked_defs: Bị can đã được liên thông chính xác sang Vụ án.
      * unlinked_defs: Bị can (đặc biệt là Mã cũ hoặc mã vụ án ngoài danh sách) chưa xác định được vụ án,
        được đánh dấu 'Xác minh sau' để kiểm tra lại trên phần mềm cũ theo yêu cầu người dùng.
    """
    def norm_s(s):
        return re.sub(r'[^a-zA-Z0-9]', '', str(s or '')).upper()
    def norm_txt(s):
        return re.sub(r'\s+', ' ', str(s or '')).strip().lower()

    if progress_cb:
        progress_cb("Đang tổng hợp danh sách vụ án...", "Khởi tạo dữ liệu...")

    if len(case_files) == 1:
        cases = [dict(c) for c in case_files[0][2]]
        case_stats = [{"fname": case_files[0][0], "total": len(cases), "matched": 0, "new": len(cases), "type": "Hồ sơ vụ án"}]
    else:
        cases, _, case_stats = reconcile_and_merge_multiple_case_lists(case_files, progress_cb)
        cases = [dict(c) for c in cases]
        for s in case_stats:
            s["type"] = "Hồ sơ vụ án"

    for c in cases:
        normalize_case_stage_and_status(c)

    case_by_code = {}
    for c in cases:
        cd = c.get("ma_vu_an")
        if cd:
            case_by_code[norm_s(cd)] = c

    def_records_all = []
    for fname, fmt, recs in defendant_files:
        for r in recs:
            r_copy = dict(r)
            r_copy["_fname"] = fname
            def_records_all.append(r_copy)

    has_blocks = any(r.get("block_idx") is not None for r in def_records_all)
    if has_blocks:
        grouped = {}
        for r in def_records_all:
            key = (r.get("_fname"), r.get("block_idx", 0))
            grouped.setdefault(key, []).append(r)
        blocks = list(grouped.values())
    else:
        grouped = {}
        blocks = []
        for r in def_records_all:
            mva = r.get("ma_vu_an")
            if mva:
                grouped.setdefault(norm_s(mva), []).append(r)
            else:
                blocks.append([r])
        blocks.extend(grouped.values())

    enriched_cases_map = {}
    linked_defs_all = []
    unlinked_defs_all = []

    for b_idx, blk in enumerate(blocks):
        if progress_cb:
            progress_cb(
                f"Đang đối chiếu & liên thông khối bị can {b_idx + 1}/{len(blocks)}...",
                f"Khối gồm {len(blk)} bị can ({blk[0].get('ten_bi_can', '')})"
            )

        target_case = None

        # 3.1 Khớp theo mã vụ án (ưu tiên)
        for d in blk:
            m_va = d.get("ma_vu_an")
            if m_va and norm_s(m_va) in case_by_code:
                target_case = case_by_code[norm_s(m_va)]
                break
            m_bc = d.get("ma_bi_can", "")
            m_match = re.match(r"^(VA_[A-Za-z0-9_\.]+?)_\d{2,3}$", m_bc)
            if m_match and norm_s(m_match.group(1)) in case_by_code:
                target_case = case_by_code[norm_s(m_match.group(1))]
                break

        # 3.2 Khớp theo CCCD hoặc Tên bị can (đặc biệt bị can đầu vụ hoặc trong tên vụ án)
        if not target_case:
            for d in blk:
                name_bc = norm_txt(d.get("ten_bi_can"))
                cccd_bc = str(d.get("so_cccd") or "").strip()
                for c in cases:
                    c_bc = norm_txt(c.get("bi_can_dau_vu") or "")
                    c_va = norm_txt(c.get("ten_vu_an") or "")
                    if cccd_bc and len(cccd_bc) >= 9 and cccd_bc in (c_bc + " " + c_va):
                        target_case = c
                        break
                    if name_bc and len(name_bc) >= 4 and (name_bc in c_bc or name_bc in c_va):
                        target_case = c
                        break
                if target_case:
                    break

        if target_case:
            c_code = norm_s(target_case.get("ma_vu_an", ""))
            enriched_cases_map[c_code] = target_case

            c_defs = []
            for idx, d in enumerate(blk, 1):
                cccd_str = f" (CCCD: {d['so_cccd']})" if d.get("so_cccd") else ""
                tt_str = f" - {d['trang_thai']}" if d.get("trang_thai") else ""
                c_defs.append(f"{idx}. {d.get('ten_bi_can', '')}{cccd_str}{tt_str}")

            existing_details = target_case.get("danh_sach_bi_can_chi_tiet", "")
            if existing_details:
                target_case["danh_sach_bi_can_chi_tiet"] = existing_details + "\n" + "\n".join(c_defs)
            else:
                target_case["danh_sach_bi_can_chi_tiet"] = "\n".join(c_defs)

            curr_total = 0
            try:
                curr_total = int(str(target_case.get("tong_bi_can") or 0).strip() or 0)
            except Exception:
                pass
            if curr_total < len(blk):
                target_case["tong_bi_can"] = len(blk)

            if blk:
                d0 = blk[0]
                lead_curr = str(target_case.get("bi_can_dau_vu") or "").lower()
                if not lead_curr or "chưa" in lead_curr:
                    cc_lead = f" - {d0.get('so_cccd')}" if d0.get("so_cccd") else ""
                    target_case["bi_can_dau_vu"] = f"{d0.get('ten_bi_can')}{cc_lead}"

            for d in blk:
                d["ma_vu_an_linked"] = target_case.get("ma_vu_an", "")
                d["ten_vu_an_linked"] = target_case.get("ten_vu_an", "")
                d["giai_doan_linked"] = target_case.get("giai_doan", "")
                d["trang_thai_va_linked"] = target_case.get("trang_thai_vu_an", "")
                d["vks_thu_ly_linked"] = target_case.get("vks_thu_ly", "")
                linked_defs_all.append(d)
        else:
            for d in blk:
                d["ma_vu_an_linked"] = "Xác minh sau"
                d["ten_vu_an_linked"] = "Chờ đối soát từ PM cũ / Bổ sung hồ sơ vụ án"
                unlinked_defs_all.append(d)

    def_stats = []
    for fname, fmt, recs in defendant_files:
        linked_in_file = sum(1 for d in linked_defs_all if d.get("_fname") == fname)
        unlinked_in_file = sum(1 for d in unlinked_defs_all if d.get("_fname") == fname)
        def_stats.append({
            "fname": fname,
            "total": len(recs),
            "matched": linked_in_file,
            "new": unlinked_in_file,
            "type": "Danh sách bị can"
        })

    all_file_stats = case_stats + def_stats
    total_enriched = len(enriched_cases_map)
    return cases, total_enriched, all_file_stats, linked_defs_all, unlinked_defs_all






def detect_data_format(raw_rows):
    """Tự động phân tích và nhận diện chính xác tuyệt đối định dạng dữ liệu đầu vào:
    - 'TIN_BAO': Danh sách chi tiết nguồn tin báo / tố giác tội phạm từ QLAHS
    - 'CASE_REPORT': Báo cáo hồ sơ vụ án chi tiết
    - 'CASE': Sổ vụ án / Sổ thụ lý vụ án / Án tạm đình chỉ / Sắp hết thời hiệu
    - 'DEFENDANT': Danh sách quản lý bị can 5 cột thô ban đầu
    """
    if not raw_rows:
        return "DEFENDANT"

    # Thu thập toàn bộ nội dung trong 25 dòng đầu tiên (không giới hạn cột)
    sample_text = " ".join(
        " ".join(str(cell or "").strip() for cell in (row[:50] if isinstance(row, (list, tuple)) else [row]))
        for row in raw_rows[:25]
    ).lower()

    # Dấu hiệu nhận biết chắc chắn: Mã tin báo thường có tiền tố 'TB_' (như TB_E01.59...)
    # Trong khi mã vụ án có tiền tố 'VA_' (như VA_E01.59...)
    has_tb_code = False
    has_va_code = False
    for row in raw_rows[:25]:
        if isinstance(row, (list, tuple)):
            for cell in row:
                cell_s = str(cell or "").strip().upper()
                if cell_s.startswith("TB_") or cell_s.startswith("TB-"):
                    has_tb_code = True
                elif cell_s.startswith("VA_") or cell_s.startswith("VA-"):
                    has_va_code = True

    # 0. Nhận diện Bảng danh sách vụ án đã chuẩn hóa phẳng (từ file kết quả đối chiếu hoặc người dùng tổng hợp/cập nhật)
    for row in raw_rows[:15]:
        if isinstance(row, (list, tuple)):
            r_str = " ".join(str(c or "").strip().lower() for c in row)
            if ("mã vụ án" in r_str or "mã va" in r_str) and (
                "đơn vị vks thụ lý" in r_str or "vks thụ lý" in r_str or 
                "mới nhất" in r_str or "số qđ khởi tố" in r_str or "nguồn dữ liệu" in r_str or "tiến độ & trạng thái" in r_str
            ):
                return "STANDARDIZED_CASE"

    # 0.1 Nhận diện Bảng danh sách Tin báo đã chuẩn hóa phẳng (từ file kết quả đối chiếu hoặc người dùng tổng hợp/cập nhật)
    for row in raw_rows[:15]:
        if isinstance(row, (list, tuple)):
            r_str = " ".join(str(c or "").strip().lower() for c in row)
            if ("mã tin báo" in r_str or "mã tb" in r_str) and (
                "đơn vị vks thụ lý" in r_str or "vks thụ lý" in r_str or 
                "mới nhất" in r_str or "người báo tin" in r_str or "nguồn dữ liệu" in r_str or 
                "thời hạn & tiến độ" in r_str or "ngày hết thời hiệu" in r_str or "cán bộ kiểm sát" in r_str
            ):
                return "STANDARDIZED_TIN_BAO"

    # 0.2 Nhận diện Sổ tin báo copy trực tiếp từ màn hình QLAHS (khối đa dòng 5 cột: STT, Thông tin tin báo, Thông tin tội danh, Trạng thái hồ sơ, Trạng thái biểu mẫu)
    for row in raw_rows[:15]:
        if isinstance(row, (list, tuple)):
            r_str = " ".join(str(c or "").strip().lower() for c in row)
            if "thông tin tin báo" in r_str and ("thông tin tội danh" in r_str or "trạng thái" in r_str or "biểu mẫu" in r_str or "giai đoạn tin báo" in r_str):
                return "TIN_BAO_RAW"
    if "thông tin tin báo" in sample_text and ("thông tin tội danh" in sample_text or "trạng thái biểu mẫu" in sample_text or "giai đoạn tin báo" in sample_text):
        return "TIN_BAO_RAW"

    # 1. Nhận diện Danh sách Tin báo về tội phạm (Ưu tiên cao nhất)
    # Lưu ý: Tin báo có thể có cột 'thời hiệu', 'tội', 'điều tra viên' nên phải kiểm tra Tin báo TRƯỚC Vụ án!
    is_tin_bao = (
        has_tb_code
        or "danh sách hồ sơ tin báo" in sample_text
        or "hồ sơ tin báo" in sample_text
        or "mã tin báo" in sample_text
        or "người bị tố giác" in sample_text
        or "thông tin xử lý tin báo" in sample_text
        or "ngày cq tổ chức tiếp nhận tin báo" in sample_text
        or "ngày hết hạn giải quyết" in sample_text
        or ("người báo tin" in sample_text and ("tin báo" in sample_text or "tố giác" in sample_text))
        or ("phân loại" in sample_text and "người báo tin" in sample_text)
        or ("tin báo" in sample_text and ("điều tra viên" in sample_text or "kiểm sát viên" in sample_text))
    )
    if is_tin_bao and not has_va_code:
        return "TIN_BAO"

    # 2. Nhận diện Danh sách Bị can / Truy nã / Đình nã (5 cột thô hoặc 10 cột chi tiết từ QLAHS)
    # Lưu ý QUAN TRỌNG: Danh sách bị can luôn chứa mã vụ án (VA_...) kèm theo từng bị can,
    # do đó PHẢI kiểm tra Danh sách Bị can / Truy nã TRƯỚC Sổ Vụ án!
    is_defendant = (
        "trạng thái hồ sơ" in sample_text
        or "đình nã" in sample_text
        or "truy nã" in sample_text
        or "mã bị can" in sample_text
        or ("thông tin hồ sơ" in sample_text and "thông tin thụ lý" not in sample_text and "hồ sơ tin báo" not in sample_text)
        or ("giai đoạn:" in sample_text and ("ngày sinh:" in sample_text or "cccd:" in sample_text or "họ tên" in sample_text) and "số thụ lý:" not in sample_text and "thông tin thụ lý" not in sample_text)
    )
    if is_defendant:
        return "DEFENDANT"

    # 3. Nhận diện Báo cáo Hồ sơ vụ án
    is_case_report = (
        "danh sách hồ sơ vụ án" in sample_text
        or ("thông tin vụ án" in sample_text and ("tội" in sample_text or "loại tội" in sample_text) and "loại thụ lý" in sample_text and "số thụ lý:" not in sample_text)
        or ("trạng thái biểu mẫu" in sample_text and "trạng thái vụ án" in sample_text and "cán bộ thụ lý" not in sample_text)
        or (has_va_code and "lưu hành nội bộ" in sample_text and "cán bộ thụ lý" not in sample_text)
    )
    if is_case_report:
        return "CASE_REPORT"

    # 4. Nhận diện Sổ vụ án / Sổ thụ lý
    if (
        has_va_code
        or "thông tin vụ án" in sample_text
        or "thông tin thụ lý" in sample_text
        or "số thụ lý:" in sample_text
        or "số qđ khởi tố:" in sample_text
        or "cán bộ thụ lý:" in sample_text
        or "thời hiệu" in sample_text
        or "loại tội" in sample_text
    ):
        return "CASE"

    return "DEFENDANT"


def parse_qlahs_tin_bao_rows(raw_rows, progress_cb=None, filename=""):
    """
    Tự động phân tích và chuẩn hóa danh sách tin báo về tội phạm từ QLAHS thành 35 cột chuẩn.
    Hỗ trợ linh hoạt cả 2 định dạng phổ biến của QLAHS:
    - Định dạng A: Chi tiết 57 cột có nhóm tiêu đề gộp
    - Định dạng B: Báo cáo hồ sơ tin báo 23 cột phẳng
    """
    if not raw_rows:
        return []

    # Tự động phát hiện và chuyển hướng nếu là dữ liệu copy thô đa dòng từ màn hình QLAHS (5 cột)
    for row in raw_rows[:15]:
        if row and any("thông tin tin báo" in str(c or "").lower() for c in row):
            return parse_qlahs_raw_tin_bao_copy_rows(raw_rows, progress_cb=progress_cb, filename=filename)

    file_context_unit = detect_file_unit(filename, raw_rows) if filename else None

    header_idx = -1
    col_map = {}

    for idx, row in enumerate(raw_rows[:15]):
        if not row:
            continue
        row_str = " ".join(str(c or "").lower() for c in row)
        if (
            "mã tin báo" in row_str
            or "danh sách hồ sơ tin báo" in row_str
            or "hồ sơ tin báo" in row_str
            or "người bị tố giác" in row_str
            or "loại tin" in row_str
            or "phân loại" in row_str
        ):
            if any(k in str(c or "").lower() for c in row for k in ("mã tin báo", "người bị tố giác", "người báo tin", "loại tin", "phân loại")):
                header_idx = idx
                break

    if header_idx != -1:
        headers = [str(c or "").strip().lower() for c in raw_rows[header_idx]]
        top_headers = [str(c or "").strip().lower() for c in raw_rows[header_idx - 1]] if header_idx > 0 else []

        for c_idx, h in enumerate(headers):
            top_h = top_headers[c_idx] if c_idx < len(top_headers) else ""
            comb = f"{top_h} {h}".strip().lower()

            if "mã tin báo" in h or "mã tin" in h:
                col_map["ma_tin_bao"] = c_idx
            elif "bị tố giác" in h or "đối tượng bị tố giác" in h:
                col_map["nguoi_bi_to_giac"] = c_idx
            elif "người báo tin" in h or "người tố giác" in h:
                col_map["nguoi_bao_tin"] = c_idx
            elif "nội dung" in h or "hành vi & hậu quả" in h or "hành vi" in h:
                col_map["noi_dung_tin_bao"] = c_idx
            elif "sắp hết hạn" in h:
                col_map["sap_het_han"] = c_idx
            elif "quá hạn" in h:
                col_map["qua_han"] = c_idx
            elif "tình trạng" in h or "kết quả giải quyết" in h or "kết quả" in h:
                col_map["ket_qua_giai_quyet"] = c_idx
            elif "ngày tạm đình chỉ" in h:
                col_map["ngay_tam_dinh_chi"] = c_idx
            elif "lý do tạm đình chỉ" in h:
                col_map["ly_do_tam_dinh_chi"] = c_idx
            elif "thời hiệu" in h or "hết thời hiệu" in h:
                col_map["ngay_het_thoi_hieu"] = c_idx
            elif "đơn vị vks thụ lý" in h or "vks thụ lý" in h or "đơn vị thụ lý" in h or "đơn vị vks" in h or h in ("đơn vị", "vks", "cơ quan thụ lý"):
                col_map["vks_thu_ly"] = c_idx
            elif "tổ chức tiếp nhận" in h or "tiếp nhận tin" in h or "đơn vị tiếp nhận" in h:
                if "ngày" in h:
                    col_map["ngay_tiep_nhan_tin"] = c_idx
                else:
                    col_map["don_vi_tiep_nhan"] = c_idx
            elif "cqđt" in h:
                col_map["ngay_cqdt_nhan"] = c_idx
            elif "điều tra viên" in h or "đtv" in h:
                col_map["dieu_tra_vien"] = c_idx
            elif "vks tiếp nhận" in h or "ngày vks nhận" in h or "vks nhận" in h:
                col_map["ngay_vks_nhan"] = c_idx
            elif "ksv phối hợp" in h or "kiểm sát viên phối hợp" in h or "phối hợp" in h or "giúp việc" in h:
                col_map["ksv_phoi_hop"] = c_idx
            elif "ksv thụ lý chính" in h or "kiểm sát viên thụ lý" in h or "ksv thụ lý" in h or "kiểm sát viên chính" in h or "ksv chính" in h or "cán bộ thụ lý" in h or "cán bộ giải quyết" in h or "cán bộ kiểm sát" in h or "kiểm sát viên" in h or h == "ksv":
                if "ksv_chinh" not in col_map:
                    col_map["ksv_chinh"] = c_idx
                else:
                    col_map["ksv_phoi_hop"] = c_idx
            elif "số tl" in h or "số thụ lý" in h:
                col_map["so_thu_ly"] = c_idx
            elif "ngày vks thụ lý" in h:
                col_map["ngay_vks_thu_ly"] = c_idx
            elif "hết hạn giải quyết" in h or "thời hạn giải quyết đến" in h or "đến ngày" in h:
                col_map["thoi_han_den_ngay"] = c_idx
            elif "thời hạn giải quyết từ" in h or "từ ngày" in h:
                col_map["thoi_han_tu_ngay"] = c_idx
            elif "phân loại" in h:
                col_map["nguon_tin"] = c_idx
                if "loai_tin" not in col_map:
                    col_map["loai_tin"] = c_idx
            elif ("loại tin" in h and "báo" in h) or h == "loại tin":
                col_map["loai_tin"] = c_idx
            elif "nguồn tin" in h:
                col_map["nguon_tin"] = c_idx
            elif "tình tiết phức tạp" in h:
                col_map["tinh_tiet_phuc_tap"] = c_idx
            elif h == "tội" or "tội danh" in h or "điều luật" in h:
                col_map["dieu_luat"] = c_idx
            elif "chương" in h:
                col_map["chuong_toi"] = c_idx
            elif "bộ luật" in h:
                col_map["bo_luat"] = c_idx
            elif "mức độ nghiêm trọng" in h:
                col_map["muc_do_nghiem_trong"] = c_idx
            elif "cccd" in h or "cmnd" in h:
                col_map["cccd_nguoi_bao"] = c_idx
            elif "số điện thoại" in h or "sđt" in h:
                col_map["sdt_nguoi_bao"] = c_idx
            elif "địa chỉ" in h and ("người" in comb or "báo tin" in comb or c_idx < 30):
                col_map["dia_chi_nguoi_bao"] = c_idx
            elif "thời gian" in h:
                col_map["thoi_gian"] = c_idx
            elif "địa điểm" in h:
                col_map["dia_diem"] = c_idx
            elif "đơn vị giải quyết" in h:
                col_map["don_vi_giai_quyet"] = c_idx
            elif "bị can khởi tố" in h or "số lượng bị can" in h:
                col_map["so_bi_can_khoi_to"] = c_idx
            elif "ghi chú" in h:
                col_map["ghi_chu"] = c_idx
            elif "mới nhất" in h or "cập nhật mới" in h or "đơn vị cập nhật" in h:
                col_map["ten_don_vi_moi_nhat"] = c_idx

        data_rows = raw_rows[header_idx + 1:]
    else:
        # Dự phòng vị trí mặc định theo mẫu 57 cột chuẩn
        col_map = {
            "loai_tin": 1, "ma_tin_bao": 2, "ngay_cqdt_nhan": 5, "ngay_vks_nhan": 6,
            "nguon_tin": 7, "tinh_tiet_phuc_tap": 8, "thoi_han_tu_ngay": 9, "thoi_han_den_ngay": 10,
            "ngay_tiep_nhan_tin": 11, "don_vi_tiep_nhan": 13, "so_thu_ly": 14, "ngay_vks_thu_ly": 15,
            "vks_thu_ly": 16, "nguoi_bao_tin": 23, "dia_chi_nguoi_bao": 25, "cccd_nguoi_bao": 26,
            "sdt_nguoi_bao": 28, "thoi_gian": 30, "dia_diem": 31, "noi_dung_tin_bao": 32,
            "dieu_luat": 35, "muc_do_nghiem_trong": 37, "don_vi_giai_quyet": 41,
            "so_bi_can_khoi_to": 43, "ket_qua_giai_quyet": 44, "ghi_chu": 46,
            "ksv_phoi_hop": 55, "ksv_chinh": 56
        }
        data_rows = raw_rows[2:]

    today_date = date.today()
    records = []

    def get_val(r, key):
        idx = col_map.get(key)
        if idx is not None and idx < len(r):
            v = r[idx]
            if v is not None:
                return str(v).strip()
        return ""

    total_rows = len(data_rows)
    for r_idx, r in enumerate(data_rows, 1):
        if progress_cb and r_idx % 500 == 0:
            progress_cb(f"Đang chuẩn hóa tin báo: {r_idx:,}/{total_rows:,} dòng...", "Tính toán hạn luật định & phân loại kết quả...")

        if not r or not any(r):
            continue
        ma_tb = get_val(r, "ma_tin_bao")
        so_tl = get_val(r, "so_thu_ly")
        noi_dung = get_val(r, "noi_dung_tin_bao")
        nguoi_bao = get_val(r, "nguoi_bao_tin")
        nguoi_bi_to = get_val(r, "nguoi_bi_to_giac")

        if not ma_tb and not so_tl and not noi_dung and not nguoi_bao:
            continue
        if ma_tb.lower() in ("mã tin báo", "mã tin", "stt"):
            continue

        rec = {}
        rec["stt"] = len(records) + 1
        rec["ma_tin_bao"] = ma_tb
        rec["loai_tin"] = get_val(r, "loai_tin") or "Mới"
        rec["nguon_tin"] = get_val(r, "nguon_tin") or "Tin báo"
        rec["tinh_tiet_phuc_tap"] = get_val(r, "tinh_tiet_phuc_tap") or "Không"

        # Người báo tin & Người bị tố giác
        rec["nguoi_bao_tin"] = get_val(r, "nguoi_bao_tin")
        rec["nguoi_bi_to_giac"] = nguoi_bi_to
        rec["cccd_nguoi_bao"] = get_val(r, "cccd_nguoi_bao")
        rec["sdt_nguoi_bao"] = get_val(r, "sdt_nguoi_bao")
        rec["dia_chi_nguoi_bao"] = get_val(r, "dia_chi_nguoi_bao")

        # Nội dung & Tội danh
        tg = get_val(r, "thoi_gian")
        dd = get_val(r, "dia_diem")
        rec["thoi_gian_dia_diem"] = f"{tg} - {dd}".strip(" -") if (tg or dd) else ""
        rec["noi_dung_tin_bao"] = noi_dung
        rec["dieu_luat"] = get_val(r, "dieu_luat")
        rec["chuong_toi"] = get_val(r, "chuong_toi")
        rec["muc_do_nghiem_trong"] = get_val(r, "muc_do_nghiem_trong")

        # Tiếp nhận & Thụ lý Kiểm sát
        raw_dv_tn = get_val(r, "don_vi_tiep_nhan")
        rec["don_vi_tiep_nhan"] = clean_vks_unit_name(raw_dv_tn) if any(k in str(raw_dv_tn).lower() for k in ["vks", "kiểm sát", "phòng", "p2", "p1", "p3", "khu vực", "thqct"]) else str(raw_dv_tn or "").strip()
        rec["ngay_tiep_nhan_tin"] = format_date_str(get_val(r, "ngay_tiep_nhan_tin"))
        rec["ngay_cqdt_nhan"] = format_date_str(get_val(r, "ngay_cqdt_nhan"))
        rec["dieu_tra_vien"] = clean_person_name(get_val(r, "dieu_tra_vien"))
        rec["ngay_vks_nhan"] = format_date_str(get_val(r, "ngay_vks_nhan"))
        rec["so_thu_ly"] = so_tl
        rec["ngay_vks_thu_ly"] = format_date_str(get_val(r, "ngay_vks_thu_ly"))

        # Chuẩn hóa tên đơn vị VKS thụ lý
        raw_vks = get_val(r, "vks_thu_ly")
        rec["vks_thu_ly"] = clean_vks_unit_name(raw_vks)
        dv_moi = get_val(r, "ten_don_vi_moi_nhat")
        if dv_moi:
            u_std = clean_vks_unit_name(dv_moi)
            if u_std and "chưa" not in u_std.lower():
                rec["vks_thu_ly"] = u_std
        if not rec.get("vks_thu_ly") or "chưa" in str(rec["vks_thu_ly"]).lower():
            u_code = parse_tin_bao_code_unit(ma_tb, file_context_unit)
            if u_code:
                rec["vks_thu_ly"] = u_code
        rec.pop("ten_don_vi_moi_nhat", None)

        # Cán bộ Kiểm sát thụ lý
        raw_ksv = get_val(r, "ksv_chinh")
        if not raw_ksv or raw_ksv.startswith("="):
            raw_ksv = get_val(r, "ksv_phoi_hop")
        if raw_ksv and not get_val(r, "ksv_phoi_hop") and any(sep in str(raw_ksv) for sep in (",", ";", "\n", "/")):
            parts = [clean_person_name(p) for p in re.split(r'[,;\n/]', str(raw_ksv)) if clean_person_name(p)]
            rec["ksv_chinh"] = parts[0] if parts else ""
            rec["ksv_phoi_hop"] = "; ".join(parts[1:]) if len(parts) > 1 else ""
        else:
            rec["ksv_chinh"] = clean_person_name(raw_ksv)
            rec["ksv_phoi_hop"] = clean_person_name(get_val(r, "ksv_phoi_hop"))
        raw_dv_gq = get_val(r, "don_vi_giai_quyet")
        rec["don_vi_giai_quyet"] = clean_vks_unit_name(raw_dv_gq) if any(k in str(raw_dv_gq).lower() for k in ["vks", "kiểm sát", "phòng", "p2", "p1", "p3", "khu vực", "thqct"]) else str(raw_dv_gq or "").strip()

        # Thời hạn & Cảnh báo tiến độ
        from_d_str = get_val(r, "thoi_han_tu_ngay")
        to_d_str = get_val(r, "thoi_han_den_ngay")
        rec["thoi_han_tu_ngay"] = format_date_str(from_d_str)
        rec["thoi_han_den_ngay"] = format_date_str(to_d_str)

        due_date = parse_date_safely(to_d_str)
        days_left = None
        if due_date:
            days_left = (due_date - today_date).days

        kq_raw = get_val(r, "ket_qua_giai_quyet") or "Đang giải quyết"
        rec["ket_qua_giai_quyet"] = kq_raw
        # Tạm đình chỉ & Thời hiệu truy cứu TNHS
        rec["ngay_tam_dinh_chi"] = format_date_str(get_val(r, "ngay_tam_dinh_chi"))
        rec["ly_do_tam_dinh_chi"] = get_val(r, "ly_do_tam_dinh_chi")

        th_raw = get_val(r, "ngay_het_thoi_hieu")
        d_th = parse_date_safely(th_raw)
        rec["ngay_het_thoi_hieu"] = format_date_str(th_raw)

        # Tính toán cảnh báo Hết thời hiệu truy cứu TNHS cho Tin báo (Group 7 & Group 6)
        if d_th:
            days_th = (d_th - today_date).days
            if days_th < 0:
                rec["so_ngay_con_lai_thoi_hieu"] = f"Đã quá {-days_th} ngày"
                rec["canh_bao_thoi_hieu"] = f"⚠️ Đã quá hạn ({-days_th} ngày)"
                rec["canh_bao_chi_tiet_thoi_hieu"] = "Đã hết thời hiệu truy cứu TNHS"
                rec["so_ngay_con_lai"] = f"Đã quá {-days_th} ngày"
                rec["muc_do_khan_cap"] = f"⚠️ Đã hết thời hiệu truy cứu TNHS (Quá {-days_th} ngày)"
            elif days_th == 0:
                rec["so_ngay_con_lai_thoi_hieu"] = "Hôm nay hết hạn"
                rec["canh_bao_thoi_hieu"] = "🔴 Hết thời hiệu hôm nay (0 ngày)"
                rec["canh_bao_chi_tiet_thoi_hieu"] = "Hết thời hiệu hôm nay"
                rec["so_ngay_con_lai"] = "Hôm nay hết hạn"
                rec["muc_do_khan_cap"] = "🔴 Hết thời hiệu hôm nay (0 ngày)"
            elif days_th <= 7:
                rec["so_ngay_con_lai_thoi_hieu"] = f"Còn {days_th} ngày"
                rec["canh_bao_thoi_hieu"] = f"🔴 Khẩn cấp (Còn {days_th} ngày)"
                rec["canh_bao_chi_tiet_thoi_hieu"] = "Sắp hết thời hiệu truy cứu TNHS"
                rec["so_ngay_con_lai"] = f"Còn {days_th} ngày"
                rec["muc_do_khan_cap"] = f"🔴 Sắp hết thời hiệu truy cứu TNHS (Còn {days_th} ngày)"
            elif days_th <= 15:
                rec["so_ngay_con_lai_thoi_hieu"] = f"Còn {days_th} ngày"
                rec["canh_bao_thoi_hieu"] = f"🟠 Rất gấp (Còn {days_th} ngày)"
                rec["canh_bao_chi_tiet_thoi_hieu"] = "Sắp hết thời hiệu truy cứu TNHS"
                rec["so_ngay_con_lai"] = f"Còn {days_th} ngày"
                rec["muc_do_khan_cap"] = f"🟠 Sắp hết thời hiệu truy cứu TNHS (Còn {days_th} ngày)"
            elif days_th <= 30:
                rec["so_ngay_con_lai_thoi_hieu"] = f"Còn {days_th} ngày"
                rec["canh_bao_thoi_hieu"] = f"🟡 Sắp hết hạn (Còn {days_th} ngày)"
                rec["canh_bao_chi_tiet_thoi_hieu"] = "Sắp hết thời hiệu truy cứu TNHS"
                rec["so_ngay_con_lai"] = f"Còn {days_th} ngày"
                rec["muc_do_khan_cap"] = f"🟡 Sắp hết thời hiệu truy cứu TNHS (Còn {days_th} ngày)"
            else:
                rec["so_ngay_con_lai_thoi_hieu"] = f"Còn {days_th} ngày"
                rec["canh_bao_thoi_hieu"] = f"🟢 Bình thường (Còn {days_th} ngày)"
                rec["canh_bao_chi_tiet_thoi_hieu"] = "Trong thời hiệu truy cứu TNHS"
                if "tạm đình chỉ" in kq_raw.lower() or rec["ngay_tam_dinh_chi"]:
                    rec["so_ngay_con_lai"] = f"Còn {days_th} ngày"
                    rec["muc_do_khan_cap"] = f"⏸ Tạm đình chỉ (Còn {days_th} ngày thời hiệu)"
        elif "không tính" in str(th_raw).lower():
            rec["so_ngay_con_lai_thoi_hieu"] = "Không tính thời hiệu"
            rec["canh_bao_thoi_hieu"] = "⚪ Không tính thời hiệu"
            rec["canh_bao_chi_tiet_thoi_hieu"] = "Không tính thời hiệu truy cứu"
        elif "tạm đình chỉ" in kq_raw.lower() or rec["ngay_tam_dinh_chi"]:
            rec["so_ngay_con_lai_thoi_hieu"] = "Chưa có ngày"
            rec["canh_bao_thoi_hieu"] = "⚠️ Tạm đình chỉ (Chưa rõ thời hiệu)"
            rec["canh_bao_chi_tiet_thoi_hieu"] = "Cần rà soát bổ sung ngày hết thời hiệu"
            rec["so_ngay_con_lai"] = "Tạm đình chỉ"
            rec["muc_do_khan_cap"] = "⏸ Tạm đình chỉ giải quyết"
        else:
            rec["so_ngay_con_lai_thoi_hieu"] = ""
            rec["canh_bao_thoi_hieu"] = ""
            rec["canh_bao_chi_tiet_thoi_hieu"] = ""

        # Xác định mức độ khẩn cấp / tiến độ giải quyết (nếu chưa được ấn định bởi thời hiệu)
        if not d_th and ("tạm đình chỉ" not in kq_raw.lower() and not rec["ngay_tam_dinh_chi"]):
            if any(k in kq_raw.lower() for k in ("khởi tố", "không khởi tố", "đã nhập", "chuyển nơi khác", "không thụ lý")):
                rec["so_ngay_con_lai"] = "Đã giải quyết"
                rec["muc_do_khan_cap"] = f"✔ Đã giải quyết ({kq_raw})"
            elif days_left is not None:
                if days_left < 0:
                    rec["so_ngay_con_lai"] = f"-{abs(days_left)} ngày"
                    rec["muc_do_khan_cap"] = f"🔴 Quá hạn giải quyết ({abs(days_left)} ngày)"
                elif days_left <= 7:
                    rec["so_ngay_con_lai"] = f"{days_left} ngày"
                    rec["muc_do_khan_cap"] = f"🟠 Khẩn cấp (Còn {days_left} ngày)"
                elif days_left <= 15:
                    rec["so_ngay_con_lai"] = f"{days_left} ngày"
                    rec["muc_do_khan_cap"] = f"🟡 Cảnh báo hạn (Còn {days_left} ngày)"
                else:
                    rec["so_ngay_con_lai"] = f"{days_left} ngày"
                    rec["muc_do_khan_cap"] = f"🟢 Trong thời hạn (Còn {days_left} ngày)"
            else:
                qua_han_txt = get_val(r, "qua_han")
                sap_het_txt = get_val(r, "sap_het_han")
                if qua_han_txt:
                    rec["so_ngay_con_lai"] = f"-{qua_han_txt}"
                    rec["muc_do_khan_cap"] = f"🔴 Quá hạn giải quyết ({qua_han_txt})"
                elif sap_het_txt:
                    rec["so_ngay_con_lai"] = sap_het_txt
                    rec["muc_do_khan_cap"] = f"🟠 Khẩn cấp (Còn {sap_het_txt})"
                else:
                    rec["so_ngay_con_lai"] = "Chưa rõ"
                    rec["muc_do_khan_cap"] = "⚪ Chưa có ngày hết hạn"

        rec["so_bi_can_khoi_to"] = get_val(r, "so_bi_can_khoi_to") or "0"
        rec["ghi_chu"] = get_val(r, "ghi_chu")

        if filename:
            rec["nguon_du_lieu"] = filename
            curr_u = str(rec.get("vks_thu_ly") or "").strip()
            is_unspecific = (
                not curr_u or 
                "chưa xác định" in curr_u.lower() or 
                "chưa rõ" in curr_u.lower() or 
                "mã cũ" in curr_u.lower() or
                curr_u.lower() in ("tp. hồ chí minh", "vksnd thành phố hồ chí minh", "vksnd tp.hcm")
            )
            if file_context_unit and is_unspecific:
                rec["vks_thu_ly"] = file_context_unit

        normalize_tin_bao_status(rec)
        records.append(rec)

    return records


def parse_qlahs_raw_tin_bao_copy_rows(raw_rows, progress_cb=None, filename=""):
    """
    Bóc tách và chuẩn hóa dữ liệu Sổ tin báo copy thô đa dòng từ màn hình QLAHS thành các bản ghi 36 cột chuẩn.
    Hỗ trợ định dạng 5 cột:
    ['STT', 'Thông tin tin báo', 'Thông tin tội danh', 'Trạng thái hồ sơ', 'Trạng thái biểu mẫu']
    Mỗi tin báo là một khối dữ liệu nhiều dòng, bắt đầu bằng STT số nguyên hoặc mã TB_...
    """
    if not raw_rows:
        return []

    file_context_unit = detect_file_unit(filename, raw_rows) if filename else None
    blocks = []
    curr = []

    for row in raw_rows:
        if not row or not any(row):
            continue
        c0 = str(row[0] or "").strip()
        c1 = str(row[1] if len(row) > 1 and row[1] is not None else "").strip()
        line_str = " ".join(str(c or "").strip().lower() for c in row)

        # Bỏ qua dòng tiêu đề cột lặp lại khi người dùng copy dán nhiều trang
        if "thông tin tin báo" in line_str or c0.lower() == "stt":
            if curr:
                blocks.append(curr)
                curr = []
            continue

        # Bắt đầu một bản ghi tin báo mới: STT là số hoặc cột 2 bắt đầu bằng TB_ / TB- hoặc mã tin báo E01 / 0102
        is_new_block = (
            bool(re.match(r'^\d+(\.0*)?$', c0))
            or c1.upper().startswith("TB_")
            or c1.upper().startswith("TB-")
            or bool(re.match(r'^(?:E01\.59|\d{5,6}\s*-)', c1))
        )
        if is_new_block:
            if curr:
                blocks.append(curr)
            curr = [row]
        else:
            if curr:
                curr.append(row)

    if curr:
        blocks.append(curr)

    today_date = date.today()
    records = []
    total_blocks = len(blocks)

    for b_idx, block in enumerate(blocks):
        if progress_cb and (b_idx + 1) % 50 == 0:
            progress_cb(f"Đang bóc tách tin báo copy thô: {b_idx + 1:,}/{total_blocks:,} tin...", "Trích xuất KSV, thời hạn & tội danh...")

        rec = {}
        rec["stt"] = len(records) + 1

        col1_texts = [str(r[1]).strip() for r in block if len(r) > 1 and r[1] is not None and str(r[1]).strip()]
        col2_texts = [str(r[2]).strip() for r in block if len(r) > 2 and r[2] is not None and str(r[2]).strip()]
        col3_texts = [str(r[3]).strip() for r in block if len(r) > 3 and r[3] is not None and str(r[3]).strip()]

        if not col1_texts:
            continue

        # Mã tin báo: ưu tiên TB_ hoặc E01.59 hoặc 0102xx, nếu không lấy dòng đầu tiên
        ma_tb = ""
        for t in col1_texts:
            if t.upper().startswith("TB_") or t.upper().startswith("TB-") or "E01.59" in t.upper() or re.match(r'^\d{5,6}\s*-\s*\d+', t):
                ma_tb = t
                break
        if not ma_tb and col1_texts:
            ma_tb = col1_texts[0]
        rec["ma_tin_bao"] = ma_tb

        noi_dung_parts = []
        raw_han_status = ""
        for line in col1_texts:
            if line == ma_tb:
                continue
            clean_l = line.replace('\xa0', ' ').strip()
            m_vks_tl = re.match(r'^Ngày VKS thụ lý:\s*(.*)', clean_l, re.I)
            m_cq_tn = re.match(r'^Ngày CQ tổ chức tiếp nhận tin báo:\s*(.*)', clean_l, re.I)
            m_han_gq = re.match(r'^Ngày hết hạn giải quyết:\s*(.*)', clean_l, re.I)
            m_pc = re.match(r'^Ngày phân công THQCT:\s*(.*)', clean_l, re.I)
            m_pl = re.match(r'^Phân loại:\s*(.*)', clean_l, re.I)
            m_kpi = re.match(r'^Số điểm KPI:\s*(.*)', clean_l, re.I)
            m_th = re.match(r'^Ngày hết thời hiệu truy cứu TNHS:\s*(.*)', clean_l, re.I)
            m_han = re.match(r'^(?:Chưa đến hạn|Quá hạn)\s*\((.*)\)', clean_l, re.I)
            m_tdc = re.match(r'^(?:Số\s*Tđc|Số\s*TĐC|TĐC)\s*[:\s]?\s*([0-9A-Za-z\-_]+)(?:[./](\d{1,2}[./]\d{1,2}[./]\d{4}))?', clean_l, re.I)

            if m_vks_tl:
                rec["ngay_vks_thu_ly"] = format_date_str(m_vks_tl.group(1).strip())
            elif m_cq_tn:
                rec["ngay_tiep_nhan_tin"] = format_date_str(m_cq_tn.group(1).strip())
            elif m_han_gq:
                rec["thoi_han_den_ngay"] = format_date_str(m_han_gq.group(1).strip())
            elif m_pc:
                rec["ngay_cqdt_nhan"] = format_date_str(m_pc.group(1).strip())
            elif m_pl:
                pl_val = m_pl.group(1).strip()
                rec["loai_tin"] = pl_val
                rec["nguon_tin"] = pl_val
            elif m_kpi:
                rec["ghi_chu"] = f"KPI: {m_kpi.group(1).strip()}"
            elif m_th:
                rec["ngay_het_thoi_hieu"] = format_date_str(m_th.group(1).strip())
            elif clean_l == "Không tính thời hiệu" or "không tính thời hiệu" in clean_l.lower():
                rec["ngay_het_thoi_hieu"] = "Không tính thời hiệu"
            elif "đã hết thời hiệu" in clean_l.lower():
                pass
            elif clean_l in ("Đã hoàn thành", "Hoàn thành"):
                rec["ket_qua_giai_quyet"] = "Đã giải quyết"
            elif m_tdc:
                rec["so_thu_ly"] = m_tdc.group(1).strip()
                if m_tdc.group(2):
                    rec["ngay_tam_dinh_chi"] = format_date_str(m_tdc.group(2).replace('.', '/'))
                rec["ket_qua_giai_quyet"] = "Tạm đình chỉ"
            elif m_han:
                raw_han_status = clean_l
            else:
                noi_dung_parts.append(clean_l)

        rec["noi_dung_tin_bao"] = " ".join(noi_dung_parts)
        if not rec.get("loai_tin"):
            rec["loai_tin"] = "Mới"
        if not rec.get("nguon_tin"):
            rec["nguon_tin"] = "Tin báo"
        rec["tinh_tiet_phuc_tap"] = "Không"

        # Trích xuất KSV & cán bộ từ nội dung tin báo
        raw_ksv_str = extract_officers_from_text(rec["noi_dung_tin_bao"])
        if raw_ksv_str:
            tokens = [t.strip().lstrip('(').strip() for t in re.split(r'[;|\n]', raw_ksv_str) if t.strip()]
            ksv_main = ""
            ksv_subs = []
            for tok in tokens:
                sub_m = re.match(r'^(ksv|pvt|cv|tp|ptp)\s*[:\s-]\s*(.+)$', tok, re.I)
                if sub_m:
                    role, name = sub_m.group(1).upper(), clean_person_name(sub_m.group(2).strip())
                    if role == "KSV" and not ksv_main:
                        ksv_main = name
                    else:
                        ksv_subs.append(f"{role}: {name}")
                else:
                    c_tok = clean_person_name(tok)
                    if c_tok:
                        ksv_subs.append(c_tok)
            if not ksv_main and ksv_subs:
                ksv_main = ksv_subs.pop(0)
            rec["ksv_chinh"] = clean_person_name(ksv_main)
            rec["ksv_phoi_hop"] = "; ".join(ksv_subs)
        else:
            rec["ksv_chinh"] = ""
            rec["ksv_phoi_hop"] = ""

        # Col 2: Chương, Tội danh, Người tố giác / Bị tố giác
        for line in col2_texts:
            clean_l = line.replace('\xa0', ' ').strip()
            m_ch = re.match(r'^Chương:\s*(.*)', clean_l, re.I)
            m_toi = re.match(r'^Tội danh:\s*(.*)', clean_l, re.I)
            m_nguoi = re.match(r'^(?:Thông tin người tố giác|Người tố giác|Người báo tin):\s*(.*)', clean_l, re.I)
            m_bito = re.match(r'^(?:Thông tin người bị tố giác|Người bị tố giác|Đối tượng bị tố giác):\s*(.*)', clean_l, re.I)
            if m_ch:
                rec["chuong_toi"] = m_ch.group(1).strip()
            elif m_toi:
                rec["dieu_luat"] = m_toi.group(1).strip()
            elif m_nguoi:
                rec["nguoi_bao_tin"] = m_nguoi.group(1).strip()
            elif m_bito:
                rec["nguoi_bi_to_giac"] = m_bito.group(1).strip()

        # Col 3: Trạng thái hồ sơ
        if not rec.get("ket_qua_giai_quyet"):
            statuses = [s for s in col3_texts if s and "giai đoạn" not in s.lower() and "trạng thái" not in s.lower()]
            kq_raw = statuses[0] if statuses else "Đang giải quyết"
            if kq_raw == "Hoàn thành":
                kq_raw = "Đã giải quyết"
            rec["ket_qua_giai_quyet"] = kq_raw

        # Đơn vị VKS thụ lý
        u = parse_tin_bao_code_unit(rec["ma_tin_bao"], file_context_unit)
        rec["vks_thu_ly"] = u

        # Thời hạn & Cảnh báo khẩn cấp
        rec["thoi_han_tu_ngay"] = rec.get("ngay_vks_thu_ly", "")
        to_d_str = rec.get("thoi_han_den_ngay", "")
        due_date = parse_date_safely(to_d_str)

        kq_val = rec["ket_qua_giai_quyet"]
        if "tạm đình chỉ" in kq_val.lower():
            th_raw = rec.get("ngay_het_thoi_hieu")
            d_th = parse_date_safely(th_raw)
            if d_th:
                days_th = (d_th - today_date).days
                if days_th < 0:
                    rec["so_ngay_con_lai_thoi_hieu"] = f"Đã quá {-days_th} ngày"
                    rec["canh_bao_thoi_hieu"] = f"⚠️ Đã quá hạn ({-days_th} ngày)"
                    rec["canh_bao_chi_tiet_thoi_hieu"] = "Đã hết thời hiệu truy cứu TNHS"
                    rec["so_ngay_con_lai"] = f"Đã quá {-days_th} ngày"
                    rec["muc_do_khan_cap"] = f"⚠️ Đã hết thời hiệu truy cứu TNHS (Quá {-days_th} ngày)"
                elif days_th <= 30:
                    rec["so_ngay_con_lai_thoi_hieu"] = f"Còn {days_th} ngày"
                    rec["canh_bao_thoi_hieu"] = f"🔴 Sắp hết hạn (Còn {days_th} ngày)"
                    rec["canh_bao_chi_tiet_thoi_hieu"] = "Sắp hết thời hiệu truy cứu TNHS"
                    rec["so_ngay_con_lai"] = f"Còn {days_th} ngày"
                    rec["muc_do_khan_cap"] = f"🔴 Sắp hết thời hiệu truy cứu TNHS (Còn {days_th} ngày)"
                else:
                    rec["so_ngay_con_lai_thoi_hieu"] = f"Còn {days_th} ngày"
                    rec["canh_bao_thoi_hieu"] = f"🟢 Bình thường (Còn {days_th} ngày)"
                    rec["canh_bao_chi_tiet_thoi_hieu"] = "Trong thời hiệu truy cứu TNHS"
                    rec["so_ngay_con_lai"] = "Tạm đình chỉ"
                    rec["muc_do_khan_cap"] = "⏸️ Tạm đình chỉ giải quyết"
            elif "không tính" in str(th_raw).lower():
                rec["so_ngay_con_lai_thoi_hieu"] = "Không tính thời hiệu"
                rec["canh_bao_thoi_hieu"] = "⚪ Không tính thời hiệu"
                rec["canh_bao_chi_tiet_thoi_hieu"] = "Không tính thời hiệu truy cứu"
                rec["so_ngay_con_lai"] = "Tạm đình chỉ"
                rec["muc_do_khan_cap"] = "⏸️ Tạm đình chỉ giải quyết"
            else:
                rec["so_ngay_con_lai_thoi_hieu"] = "Chưa rõ"
                rec["canh_bao_thoi_hieu"] = "⚠️ Tạm đình chỉ (Chưa rõ thời hiệu)"
                rec["canh_bao_chi_tiet_thoi_hieu"] = "Cần rà soát bổ sung ngày hết thời hiệu"
                rec["so_ngay_con_lai"] = "Tạm đình chỉ"
                rec["muc_do_khan_cap"] = "⏸️ Tạm đình chỉ giải quyết"
        elif "đã giải quyết" in kq_val.lower() or "hoàn thành" in kq_val.lower():
            rec["so_ngay_con_lai"] = "Đã giải quyết"
            rec["muc_do_khan_cap"] = "✔ Đã giải quyết (Hoàn thành)"
        elif due_date:
            days_left = (due_date - today_date).days
            if days_left < 0:
                rec["so_ngay_con_lai"] = f"-{abs(days_left)} ngày"
                rec["muc_do_khan_cap"] = f"🔴 Quá hạn giải quyết ({abs(days_left)} ngày)"
            elif days_left <= 7:
                rec["so_ngay_con_lai"] = f"{days_left} ngày"
                rec["muc_do_khan_cap"] = f"🟠 Khẩn cấp (Còn {days_left} ngày)"
            elif days_left <= 15:
                rec["so_ngay_con_lai"] = f"{days_left} ngày"
                rec["muc_do_khan_cap"] = f"🟡 Cảnh báo hạn (Còn {days_left} ngày)"
            else:
                rec["so_ngay_con_lai"] = f"{days_left} ngày"
                rec["muc_do_khan_cap"] = f"🟢 Trong thời hạn (Còn {days_left} ngày)"
        elif raw_han_status:
            rec["so_ngay_con_lai"] = raw_han_status
            rec["muc_do_khan_cap"] = f"🟢 {raw_han_status}"
        else:
            rec["so_ngay_con_lai"] = "Chưa rõ"
            rec["muc_do_khan_cap"] = "⚪ Chưa có ngày hết hạn"

        rec["so_bi_can_khoi_to"] = "0"
        if filename:
            rec["nguon_du_lieu"] = filename

        normalize_tin_bao_status(rec)
        records.append(rec)

    return records


def build_pivot_tin_bao_summary(records):
    """Xây dựng bảng tổng hợp Pivot: Đơn vị VKS thụ lý × Kết quả giải quyết tin báo."""
    common_order = ["Khởi tố", "Không khởi tố", "Tạm đình chỉ", "Đang giải quyết", "Chuyển nơi khác", "Không thụ lý", "Đã nhập"]
    results_found = set()
    units_found = set()
    for r in records:
        kq = r.get("ket_qua_giai_quyet", "").strip() or "Đang giải quyết"
        results_found.add(kq)
        u = r.get("vks_thu_ly", "").strip() or "Chưa rõ đơn vị VKS"
        units_found.add(u)

    sorted_results = [x for x in common_order if x in results_found]
    for x in sorted(results_found):
        if x not in sorted_results:
            sorted_results.append(x)

    matrix = {}
    for u in units_found:
        matrix[u] = {k: 0 for k in sorted_results}
        matrix[u]["_total"] = 0

    for r in records:
        u = r.get("vks_thu_ly", "").strip() or "Chưa rõ đơn vị VKS"
        kq = r.get("ket_qua_giai_quyet", "").strip() or "Đang giải quyết"
        if u in matrix and kq in matrix[u]:
            matrix[u][kq] += 1
            matrix[u]["_total"] += 1

    sorted_units = sorted(matrix.keys(), key=lambda x: matrix[x]["_total"], reverse=True)
    return matrix, sorted_results, sorted_units


def set_range_border(ws, cell_range, border=None):
    """Thiết lập viền ô cho dải ô merge trong openpyxl."""
    if border is None:
        thin_side = Side(border_style="thin", color="CBD5E1")
        border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    for row in ws[cell_range]:
        for cell in row:
            cell.border = border


def create_tin_bao_thoi_hieu_pivot_sheet(wb, records, title="Pivot_Thoi_Hieu_Tin_Bao", index=1, report_title=None):
    """
    Tạo Sheet Pivot Table cho Tin báo theo đúng cấu trúc:
    - Hàng (Rows): Đơn vị VKS thụ lý (vks_thu_ly)
    - Cột (Columns): Cảnh báo chi tiết thời hiệu (canh_bao_chi_tiet_thoi_hieu)
    - Giá trị (Values): Count of Mã tin báo
    - Tổng cộng & Tỷ lệ %
    """
    if not records:
        return None

    col_order_pref = [
        "Cần rà soát bổ sung ngày hết thời hiệu",
        "Đã hết thời hiệu truy cứu TNHS",
        "Hết thời hiệu hôm nay",
        "Không tính thời hiệu truy cứu",
        "Sắp hết thời hiệu truy cứu TNHS",
        "Trong thời hiệu truy cứu TNHS",
    ]

    matrix = {}
    found_cols = set()
    for r in records:
        u = str(r.get("vks_thu_ly") or "").strip() or "Chưa rõ đơn vị VKS"
        cb = str(r.get("canh_bao_chi_tiet_thoi_hieu") or "").strip()
        if not cb:
            cb = "(Trống)"
        found_cols.add(cb)
        if u not in matrix:
            matrix[u] = {"total": 0, "cols": {}}
        matrix[u]["total"] += 1
        matrix[u]["cols"][cb] = matrix[u]["cols"].get(cb, 0) + 1

    active_cols = [c for c in col_order_pref if c in found_cols]
    for c in sorted(found_cols):
        if c != "(Trống)" and c not in active_cols:
            active_cols.append(c)
    if "(Trống)" in found_cols:
        active_cols.append("(Trống)")

    sorted_units = sorted(matrix.keys())
    tot_all = len(records)

    ws = wb.create_sheet(title=title, index=index)
    ws.views.sheetView[0].showGridLines = True

    NAVY = "07172C"
    NAVY_LIGHT = "0F2942"
    WHITE = "FFFFFF"
    BORDER_CLR = "CBD5E1"
    ZEBRA_BG = "F8FAFC"

    thin_border = Border(
        left=Side(style="thin", color=BORDER_CLR),
        right=Side(style="thin", color=BORDER_CLR),
        top=Side(style="thin", color=BORDER_CLR),
        bottom=Side(style="thin", color=BORDER_CLR),
    )
    double_bottom_border = Border(
        left=Side(style="thin", color=BORDER_CLR),
        right=Side(style="thin", color=BORDER_CLR),
        top=Side(style="thin", color=BORDER_CLR),
        bottom=Side(style="double", color="0F172A"),
    )

    col_gt = len(active_cols) + 2
    col_pct = col_gt + 1
    last_col_letter = get_column_letter(col_pct)

    now = datetime.now()
    if not report_title:
        time_str = f"{now.hour}H{now.minute:02d} NGÀY {now.strftime('%d/%m/%Y')}"
        report_title = f"DANH SÁCH VỤ VIỆC CẢNH BÁO THỜI HIỆU TRUY CỨU TNHS CẬP NHẬT LÚC {time_str}"

    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 22
    ws.merge_cells(f"A2:{last_col_letter}3")
    c_title = ws.cell(row=2, column=1, value=report_title)
    c_title.font = Font(name="Calibri", size=13, bold=True, color="000000")
    c_title.alignment = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[4].height = 30
    c_rl = ws.cell(row=4, column=1, value="ĐƠN VỊ")
    c_rl.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    c_rl.fill = PatternFill(start_color=NAVY_LIGHT, end_color=NAVY_LIGHT, fill_type="solid")
    c_rl.alignment = Alignment(horizontal="left", vertical="center")
    c_rl.border = thin_border

    for idx_c, col_name in enumerate(active_cols, 2):
        c = ws.cell(row=4, column=idx_c, value=col_name)
        c.font = Font(name="Calibri", size=10, bold=True, color=WHITE)
        c.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = thin_border

    c_gt = ws.cell(row=4, column=col_gt, value="TỔNG SỐ TIN")
    c_gt.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    c_gt.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    c_gt.alignment = Alignment(horizontal="center", vertical="center")
    c_gt.border = thin_border

    c_pct_h = ws.cell(row=4, column=col_pct, value="Tỷ lệ %")
    c_pct_h.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    c_pct_h.fill = PatternFill(start_color="047857", end_color="047857", fill_type="solid")
    c_pct_h.alignment = Alignment(horizontal="center", vertical="center")
    c_pct_h.border = thin_border

    cur_row = 5
    for u_idx, u in enumerate(sorted_units, 1):
        ws.row_dimensions[cur_row].height = 22
        st = matrix[u]
        pct = (st["total"] / tot_all * 100) if tot_all > 0 else 0.0
        bg_c = ZEBRA_BG if u_idx % 2 == 0 else WHITE

        c_u = ws.cell(row=cur_row, column=1, value=u)
        c_u.font = Font(name="Calibri", size=11)
        c_u.alignment = Alignment(horizontal="left", vertical="center")
        c_u.fill = PatternFill(start_color=bg_c, end_color=bg_c, fill_type="solid")
        c_u.border = thin_border

        for idx_c, col_name in enumerate(active_cols, 2):
            v = st["cols"].get(col_name, 0)
            c = ws.cell(row=cur_row, column=idx_c, value=v if v > 0 else "")
            c.font = Font(name="Calibri", size=11, bold=(v > 0))
            c.alignment = Alignment(horizontal="right", vertical="center")
            c.fill = PatternFill(start_color=bg_c, end_color=bg_c, fill_type="solid")
            c.border = thin_border

        c_tot = ws.cell(row=cur_row, column=col_gt, value=st["total"])
        c_tot.font = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
        c_tot.alignment = Alignment(horizontal="right", vertical="center")
        c_tot.fill = PatternFill(start_color=bg_c, end_color=bg_c, fill_type="solid")
        c_tot.border = thin_border

        c_p = ws.cell(row=cur_row, column=col_pct, value=f"{pct:.1f}%")
        c_p.font = Font(name="Calibri", size=11)
        c_p.alignment = Alignment(horizontal="right", vertical="center")
        c_p.fill = PatternFill(start_color=bg_c, end_color=bg_c, fill_type="solid")
        c_p.border = thin_border

        cur_row += 1

    ws.row_dimensions[cur_row].height = 26
    c_gt_label = ws.cell(row=cur_row, column=1, value="TỔNG CỘNG")
    c_gt_label.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    c_gt_label.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    c_gt_label.alignment = Alignment(horizontal="left", vertical="center")
    c_gt_label.border = double_bottom_border

    for idx_c, col_name in enumerate(active_cols, 2):
        tot_c = sum(matrix[u]["cols"].get(col_name, 0) for u in sorted_units)
        c = ws.cell(row=cur_row, column=idx_c, value=tot_c)
        c.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
        c.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
        c.alignment = Alignment(horizontal="right", vertical="center")
        c.border = double_bottom_border

    c_all_gt = ws.cell(row=cur_row, column=col_gt, value=tot_all)
    c_all_gt.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    c_all_gt.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    c_all_gt.alignment = Alignment(horizontal="right", vertical="center")
    c_all_gt.border = double_bottom_border

    c_all_pct = ws.cell(row=cur_row, column=col_pct, value="100.0%")
    c_all_pct.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    c_all_pct.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    c_all_pct.alignment = Alignment(horizontal="right", vertical="center")
    c_all_pct.border = double_bottom_border

    ws.column_dimensions["A"].width = 36
    for idx_c in range(2, col_gt):
        ws.column_dimensions[get_column_letter(idx_c)].width = 24
    ws.column_dimensions[get_column_letter(col_gt)].width = 16
    ws.column_dimensions[get_column_letter(col_pct)].width = 13

    ws.freeze_panes = "B5"
    return ws


def try_add_native_excel_pivot_tin_bao(out_filepath, sheet_name="Pivot_Thoi_Hieu_Tin_Bao"):
    """
    Nâng cấp sheet Pivot_Thoi_Hieu_Tin_Bao thành PivotTable gốc của Microsoft Excel (qua COM)
    để người dùng giữ nguyên bảng PivotTable Fields (kéo thả, lọc, double-click xem chi tiết).
    Nếu máy không cài Microsoft Excel hoặc COM lỗi, tự động giữ nguyên bảng thống kê tĩnh.
    """
    try:
        import win32com.client
    except ImportError:
        return False

    excel = None
    wb_com = None
    try:
        abs_path = os.path.abspath(out_filepath)
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb_com = excel.Workbooks.Open(abs_path)
        ws_data = wb_com.Sheets("Danh_Sach_Tin_Bao_Chuan_Hoa")

        last_row = ws_data.UsedRange.Row + ws_data.UsedRange.Rows.Count - 1
        last_col = ws_data.Cells(5, ws_data.Columns.Count).End(-4159).Column  # xlToLeft
        if last_row <= 5 or last_col <= 1:
            wb_com.Close(SaveChanges=False)
            return False

        src_range = ws_data.Range(ws_data.Cells(5, 1), ws_data.Cells(last_row, last_col))

        # Xóa sheet tĩnh cũ nếu đã tồn tại để thay bằng PivotTable gốc
        for sh in wb_com.Sheets:
            if sh.Name == sheet_name:
                sh.Delete()
                break

        ws_pivot = wb_com.Sheets.Add(After=ws_data)
        ws_pivot.Name = sheet_name
        ws_pivot = wb_com.Sheets(sheet_name)

        pc = wb_com.PivotCaches().Create(SourceType=1, SourceData=src_range)  # xlDatabase = 1
        pt = pc.CreatePivotTable(TableDestination=ws_pivot.Range("A3"), TableName="PivotThoiHieuTinBao")

        # Rows: Đơn vị VKS thụ lý
        pf_row = pt.PivotFields("Đơn vị VKS thụ lý")
        pf_row.Orientation = 1  # xlRowField

        # Columns: Cảnh báo chi tiết
        pf_col = pt.PivotFields("Cảnh báo chi tiết")
        pf_col.Orientation = 2  # xlColumnField

        # Values: Số lượng vụ việc (Count of Mã tin báo)
        pt.AddDataField(pt.PivotFields("Mã tin báo"), "Số lượng vụ việc", -4112)  # xlCount

        # Đổi tên nhãn mặc định sang tiếng Việt chuẩn mà vẫn giữ nguyên PivotTable Fields
        try:
            pt.CompactLayoutRowHeader = "ĐƠN VỊ"
            pt.CompactLayoutColumnHeader = "CẢNH BÁO THỜI HIỆU"
            pt.GrandTotalName = "TỔNG CỘNG"
        except Exception:
            pass

        # Định dạng gọn gàng, dễ đọc
        try:
            pt.TableStyle2 = "PivotStyleMedium2"
        except Exception:
            pass

        # Thêm tiêu đề danh sách ở dòng 1-2 (gộp ô theo đúng bề rộng bảng Pivot)
        try:
            pt_cols = max(pt.TableRange2.Columns.Count, 5)
            now = datetime.now()
            time_str = f"{now.hour}H{now.minute:02d} NGÀY {now.strftime('%d/%m/%Y')}"
            banner_txt = f"DANH SÁCH VỤ VIỆC CẢNH BÁO THỜI HIỆU TRUY CỨU TNHS CẬP NHẬT LÚC {time_str}"
            title_rng = ws_pivot.Range(ws_pivot.Cells(1, 1), ws_pivot.Cells(2, pt_cols))
            title_rng.Merge()
            title_rng.Value = banner_txt
            title_rng.Font.Name = "Calibri"
            title_rng.Font.Size = 13
            title_rng.Font.Bold = True
            title_rng.HorizontalAlignment = -4108  # xlCenter
            title_rng.VerticalAlignment = -4108    # xlCenter
            ws_pivot.Rows(1).RowHeight = 20
            ws_pivot.Rows(2).RowHeight = 22
        except Exception:
            pass

        ws_pivot.Columns.AutoFit()
        ws_pivot.Activate()

        wb_com.Save()
        wb_com.Close(SaveChanges=False)
        return True
    except Exception:
        try:
            if wb_com is not None:
                wb_com.Close(SaveChanges=False)
        except Exception:
            pass
        return False
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def export_tin_bao_to_excel(records, out_filepath):
    """Xuất file Excel chuẩn hóa Danh sách Tin báo về tội phạm (hỗ trợ cả 35 cột đơn file và 36 cột hợp nhất đa nguồn)."""
    wb_out = openpyxl.Workbook()

    ws1 = wb_out.active
    ws1.title = "Danh_Sach_Tin_Bao_Chuan_Hoa"
    ws1.views.sheetView[0].showGridLines = True

    has_sources = bool(records and any(r.get("nguon_du_lieu") for r in records))

    cols_config_base = [
        ("stt", "STT", 7, "center", "1E3A8A"),
        ("ma_tin_bao", "Mã tin báo", 18, "center", "1E3A8A"),
        ("loai_tin", "Loại tin", 12, "center", "1E3A8A"),
        ("nguon_tin", "Phân loại nguồn tin", 16, "center", "1E3A8A"),
        ("tinh_tiet_phuc_tap", "Phức tạp", 10, "center", "1E3A8A"),

        ("nguoi_bao_tin", "Người / Cơ quan báo tin", 24, "left", "1E40AF"),
        ("nguoi_bi_to_giac", "Người / Đối tượng bị tố giác", 24, "left", "1E40AF"),
        ("cccd_nguoi_bao", "CCCD / Định danh", 16, "center", "1E40AF"),
        ("sdt_nguoi_bao", "Số điện thoại", 14, "center", "1E40AF"),
        ("dia_chi_nguoi_bao", "Địa chỉ người báo", 25, "left", "1E40AF"),

        ("thoi_gian_dia_diem", "Thời gian & Địa điểm", 22, "left", "0F766E"),
        ("noi_dung_tin_bao", "Nội dung tóm tắt nguồn tin", 38, "left", "0F766E"),
        ("dieu_luat", "Thông tin điều luật / Tội danh", 30, "left", "0F766E"),
        ("chuong_toi", "Chương BLHS", 24, "left", "0F766E"),
        ("muc_do_nghiem_trong", "Mức độ tội", 16, "center", "0F766E"),

        ("don_vi_tiep_nhan", "Đơn vị tiếp nhận ban đầu", 24, "left", "0369A1"),
        ("ngay_tiep_nhan_tin", "Ngày tiếp nhận tin", 14, "center", "0369A1"),
        ("ngay_cqdt_nhan", "Ngày CQĐT thụ lý", 14, "center", "0369A1"),
        ("dieu_tra_vien", "Điều tra viên (ĐTV)", 20, "left", "0369A1"),
        ("ngay_vks_nhan", "Ngày VKS tiếp nhận", 14, "center", "0369A1"),
        ("so_thu_ly", "Số thụ lý (Số TL)", 14, "center", "0369A1"),
        ("ngay_vks_thu_ly", "Ngày VKS thụ lý", 14, "center", "0369A1"),
        ("vks_thu_ly", "Đơn vị VKS thụ lý", 24, "left", "0369A1"),

        ("ksv_chinh", "Kiểm sát viên thụ lý chính", 22, "left", "5B21B6"),
        ("ksv_phoi_hop", "KSV phối hợp / giúp việc", 20, "left", "5B21B6"),
        ("don_vi_giai_quyet", "Đơn vị giải quyết", 20, "left", "5B21B6"),

        ("thoi_han_tu_ngay", "Thời hạn từ ngày", 14, "center", "991B1B"),
        ("thoi_han_den_ngay", "Thời hạn đến ngày", 14, "center", "991B1B"),
        ("so_ngay_con_lai", "Số ngày còn lại", 15, "center", "991B1B"),
        ("muc_do_khan_cap", "Cảnh báo tiến độ / Thời hạn", 25, "left", "991B1B"),
        ("ket_qua_giai_quyet", "Tình trạng / Kết quả giải quyết", 18, "center", "991B1B"),

        ("ngay_tam_dinh_chi", "Ngày tạm đình chỉ", 14, "center", "991B1B"),
        ("ly_do_tam_dinh_chi", "Lý do tạm đình chỉ", 24, "left", "334155"),
        ("ngay_het_thoi_hieu", "Ngày hết thời hiệu truy cứu", 16, "center", "991B1B"),
        ("so_ngay_con_lai_thoi_hieu", "Số ngày còn lại thời hiệu", 15, "center", "991B1B"),
        ("canh_bao_thoi_hieu", "Cảnh báo tiến độ / Khẩn cấp", 25, "center", "991B1B"),
        ("canh_bao_chi_tiet_thoi_hieu", "Cảnh báo chi tiết", 25, "center", "991B1B"),
        ("ghi_chu", "Ghi chú", 20, "left", "334155"),
    ]

    if has_sources:
        cols_config = [
            ("stt", "STT", 7, "center", "1E3A8A"),
            ("nguon_du_lieu", "Nguồn dữ liệu đối chiếu", 30, "left", "1E3A8A"),
        ] + [c for c in cols_config_base if c[0] != "stt"]
        groups_config = [
            ("A4:F4", "1. NHẬN DIỆN, ĐỐI CHIẾU & NGUỒN TIN", "1E3A8A", "E0E7FF"),
            ("G4:K4", "2. NGƯỜI BÁO TIN & ĐỐI TƯỢNG BỊ TỐ GIÁC", "1E40AF", "DBEAFE"),
            ("L4:P4", "3. NỘI DUNG NGUỒN TIN & TỘI DANH", "0F766E", "CCFBF1"),
            ("Q4:X4", "4. TIẾP NHẬN & THỤ LÝ KIỂM SÁT", "0369A1", "E0F2FE"),
            ("Y4:AA4", "5. CÁN BỘ KIỂM SÁT THỤ LÝ", "5B21B6", "EDE9FE"),
            ("AB4:AF4", "6. THỜI HẠN & TIẾN ĐỘ GIẢI QUYẾT", "991B1B", "FEE2E2"),
            ("AG4:AM4", "7. TẠM ĐÌNH CHỈ & THỜI HIỆU TRUY CỨU TNHS (ĐẶC BIỆT CHÚ Ý)", "991B1B", "FEE2E2"),
        ]
        freeze_cell = "D6"
        banner_sub = f"BẢNG HỢP NHẤT & ĐỐI CHIẾU ĐA NGUỒN DANH SÁCH TIN BÁO TỐ GIÁC TỘI PHẠM QLAHS (Tổng hợp {len(records):,} tin báo)"
    else:
        cols_config = cols_config_base
        groups_config = [
            ("A4:E4", "1. NHẬN DIỆN & NGUỒN TIN", "1E3A8A", "E0E7FF"),
            ("F4:J4", "2. NGƯỜI BÁO TIN & ĐỐI TƯỢNG BỊ TỐ GIÁC", "1E40AF", "DBEAFE"),
            ("K4:O4", "3. NỘI DUNG NGUỒN TIN & TỘI DANH", "0F766E", "CCFBF1"),
            ("P4:W4", "4. TIẾP NHẬN & THỤ LÝ KIỂM SÁT", "0369A1", "E0F2FE"),
            ("X4:Z4", "5. CÁN BỘ KIỂM SÁT THỤ LÝ", "5B21B6", "EDE9FE"),
            ("AA4:AE4", "6. THỜI HẠN & TIẾN ĐỘ GIẢI QUYẾT", "991B1B", "FEE2E2"),
            ("AF4:AL4", "7. TẠM ĐÌNH CHỈ & THỜI HIỆU TRUY CỨU TNHS (ĐẶC BIỆT CHÚ Ý)", "991B1B", "FEE2E2"),
        ]
        freeze_cell = "C6"
        banner_sub = f"BẢNG TỔNG HỢP & CHUẨN HÓA DANH SÁCH TIN BÁO TỐ GIÁC TỘI PHẠM QLAHS (Tổng hợp {len(records):,} tin báo)"

    max_col_letter = get_column_letter(len(cols_config))

    ws1.row_dimensions[1].height = 28
    ws1.row_dimensions[2].height = 20
    ws1.row_dimensions[3].height = 10
    ws1.row_dimensions[4].height = 26
    ws1.row_dimensions[5].height = 32

    # Banner Header
    ws1.merge_cells(f"A1:{max_col_letter}1")
    cell_b1 = ws1["A1"]
    cell_b1.value = "VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH"
    cell_b1.font = Font(name="Segoe UI", size=15, bold=True, color="FFFFFF")
    cell_b1.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    cell_b1.alignment = Alignment(horizontal="center", vertical="center")

    ws1.merge_cells(f"A2:{max_col_letter}2")
    cell_b2 = ws1["A2"]
    time_stamp_full = datetime.now().strftime("%d/%m/%Y %H:%M")
    cell_b2.value = f"{banner_sub} - Xuất ngày: {time_stamp_full}"
    cell_b2.font = Font(name="Segoe UI", size=10, italic=True, color="E0E7FF")
    cell_b2.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    cell_b2.alignment = Alignment(horizontal="center", vertical="center")

    # Dòng nhóm 4
    for rng, gtitle, bg_c, fg_c in groups_config:
        ws1.merge_cells(rng)
        first_c = ws1[rng.split(":")[0]]
        first_c.value = gtitle
        first_c.font = Font(name="Segoe UI", size=10, bold=True, color=bg_c)
        first_c.fill = PatternFill(start_color=fg_c, end_color=fg_c, fill_type="solid")
        first_c.alignment = Alignment(horizontal="center", vertical="center")
        set_range_border(ws1, rng)

    # Dòng 5: Chi tiết các cột
    for c_idx, (_, col_title, width, _, bg_c) in enumerate(cols_config, start=1):
        c = ws1.cell(row=5, column=c_idx)
        c.value = col_title
        c.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color=bg_c, end_color=bg_c, fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        col_letter = get_column_letter(c_idx)
        ws1.column_dimensions[col_letter].width = width

    # Ghi dữ liệu dòng 6 trở đi
    thin_side = Side(border_style="thin", color="CBD5E1")
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    font_data = Font(name="Segoe UI", size=9)
    font_data_bold = Font(name="Segoe UI", size=9, bold=True)

    fill_even = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_odd = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    fill_overdue = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fill_urgent = PatternFill(start_color="FFEDD5", end_color="FFEDD5", fill_type="solid")
    fill_warning = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")
    fill_safe = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")

    for r_idx, rec in enumerate(records, start=6):
        ws1.row_dimensions[r_idx].height = 22
        base_fill = fill_even if (r_idx % 2 == 0) else fill_odd
        urg = rec.get("muc_do_khan_cap", "")

        for c_idx, (col_key, _, _, align, _) in enumerate(cols_config, start=1):
            c = ws1.cell(row=r_idx, column=c_idx)
            val = rec.get(col_key, "")
            c.border = cell_border
            c.font = font_data
            c.alignment = Alignment(horizontal=align, vertical="center")

            # 1. Định dạng Ngày tháng chuẩn Excel DD/MM/YYYY cho tất cả các cột ngày
            if col_key in ("ngay_tiep_nhan_tin", "ngay_cqdt_nhan", "ngay_vks_nhan", "ngay_vks_thu_ly",
                           "thoi_han_tu_ngay", "thoi_han_den_ngay", "ngay_tam_dinh_chi", "ngay_het_thoi_hieu"):
                dt_obj = parse_date_safely(val)
                if dt_obj:
                    c.value = dt_obj
                    c.number_format = "DD/MM/YYYY"
                else:
                    c.value = str(val or "")
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.fill = base_fill

            # 2. Cột Cảnh báo tiến độ / Thời hạn giải quyết (Group 6)
            elif col_key == "muc_do_khan_cap":
                c.value = str(val or "")
                c.font = font_data_bold
                c.alignment = Alignment(horizontal="center", vertical="center")
                val_s = str(val or "")
                if "🔴" in val_s or "quá hạn" in val_s.lower():
                    c.fill = fill_overdue
                    c.font = Font(name="Segoe UI", size=9, bold=True, color="991B1B")
                elif "🟠" in val_s or "khẩn cấp" in val_s.lower():
                    c.fill = fill_urgent
                    c.font = Font(name="Segoe UI", size=9, bold=True, color="C2410C")
                elif "🟡" in val_s or "cảnh báo" in val_s.lower():
                    c.fill = fill_warning
                    c.font = Font(name="Segoe UI", size=9, bold=True, color="854D0E")
                elif "🟢" in val_s:
                    c.fill = fill_safe
                    c.font = Font(name="Segoe UI", size=9, bold=True, color="166534")
                else:
                    c.fill = base_fill

            elif col_key == "so_ngay_con_lai":
                c.value = str(val or "")
                c.font = font_data_bold
                c.alignment = Alignment(horizontal="center", vertical="center")
                if "-" in str(val) or "quá" in str(val).lower():
                    c.font = Font(name="Segoe UI", size=9, bold=True, color="DC2626")
                c.fill = base_fill

            # 3. Cột Số ngày còn lại thời hiệu TNHS (Group 7)
            elif col_key == "so_ngay_con_lai_thoi_hieu":
                c.value = str(val or "")
                c.font = font_data_bold
                c.alignment = Alignment(horizontal="center", vertical="center")
                val_s = str(val or "").lower()
                if "quá" in val_s or "-" in val_s:
                    c.font = Font(name="Segoe UI", size=9, bold=True, color="991B1B")
                    c.fill = fill_overdue
                elif "còn" in val_s:
                    m_d = re.search(r'\b(\d+)\b', val_s)
                    if m_d and int(m_d.group(1)) <= 30:
                        c.fill = fill_warning
                        c.font = Font(name="Segoe UI", size=9, bold=True, color="854D0E")
                    else:
                        c.fill = base_fill
                else:
                    c.fill = base_fill

            # 4. Cột Cảnh báo thời hiệu & Cảnh báo chi tiết (Group 7)
            elif col_key in ("canh_bao_thoi_hieu", "canh_bao_chi_tiet_thoi_hieu"):
                c.value = str(val or "")
                c.alignment = Alignment(horizontal="center", vertical="center")
                val_s = str(val or "")
                if "🔴" in val_s or "quá hạn" in val_s.lower() or "đã hết" in val_s.lower():
                    c.fill = fill_overdue
                    c.font = Font(name="Segoe UI", size=9, bold=True, color="991B1B")
                elif "🟠" in val_s or "khẩn cấp" in val_s.lower() or "rất gấp" in val_s.lower():
                    c.fill = fill_urgent
                    c.font = Font(name="Segoe UI", size=9, bold=True, color="C2410C")
                elif "🟡" in val_s or "sắp hết" in val_s.lower():
                    c.fill = fill_warning
                    c.font = Font(name="Segoe UI", size=9, bold=True, color="854D0E")
                elif "🟢" in val_s or "còn" in val_s.lower() or "trong thời hiệu" in val_s.lower():
                    c.fill = fill_safe
                    c.font = Font(name="Segoe UI", size=9, bold=True, color="166534")
                elif "⚠️" in val_s:
                    c.fill = fill_urgent
                    c.font = Font(name="Segoe UI", size=9, bold=True, color="C2410C")
                else:
                    c.fill = base_fill

            else:
                c.value = str(val or "")
                c.fill = base_fill

            if col_key in ("cccd_nguoi_bao", "sdt_nguoi_bao", "so_thu_ly"):
                c.number_format = "@"

    ws1.freeze_panes = freeze_cell
    ws1.auto_filter.ref = f"A5:{max_col_letter}{len(records) + 5}"

    # Sheet 2: Báo cáo Tiến độ giải quyết Tin báo & Thống kê KSV
    ws2 = wb_out.create_sheet(title="Bao_Cao_Tien_Do_Tin_Bao")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:G1")
    t2 = ws2["A1"]
    t2.value = "BÁO CÁO TIẾN ĐỘ GIẢI QUYẾT & PHÂN CÔNG KIỂM SÁT VIÊN THỤ LÝ TIN BÁO"
    t2.font = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
    t2.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    t2.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 28

    total_tb = len(records)
    overdue = sum(1 for r in records if "🔴" in r.get("muc_do_khan_cap", ""))
    urgent = sum(1 for r in records if "🟠" in r.get("muc_do_khan_cap", ""))
    warning = sum(1 for r in records if "🟡" in r.get("muc_do_khan_cap", ""))
    safe = sum(1 for r in records if "🟢" in r.get("muc_do_khan_cap", ""))
    solved = sum(1 for r in records if "✔" in r.get("muc_do_khan_cap", ""))
    suspended = sum(1 for r in records if "⏸" in r.get("muc_do_khan_cap", ""))

    kpi_rows = [
        ("1", "Tổng số tin báo tội phạm thụ lý trên hệ thống", total_tb, "100.0%", "Toàn bộ nguồn tin đã số hóa"),
        ("2", "Đã giải quyết (Khởi tố / Không khởi tố / Chuyển nơi khác)", solved, f"{(solved/total_tb*100) if total_tb else 0:.1f}%", "Đã hoàn thành tố tụng"),
        ("3", "Tạm đình chỉ giải quyết nguồn tin", suspended, f"{(suspended/total_tb*100) if total_tb else 0:.1f}%", "Chờ kết quả giám định / xác minh"),
        ("4", "Đang trong thời hạn giải quyết an toàn (> 15 ngày)", safe, f"{(safe/total_tb*100) if total_tb else 0:.1f}%", "Tiến độ trong tầm kiểm soát"),
        ("5", "Cảnh báo thời hạn giải quyết (Còn 8 - 15 ngày)", warning, f"{(warning/total_tb*100) if total_tb else 0:.1f}%", "Cần đẩy nhanh tiến độ"),
        ("6", "Khẩn cấp sắp hết thời hạn giải quyết (Còn ≤ 7 ngày)", urgent, f"{(urgent/total_tb*100) if total_tb else 0:.1f}%", "Yêu cầu ra quyết định khẩn"),
        ("7", "QUÁ HẠN GIẢI QUYẾT LUẬT ĐỊNH (Cần chỉ đạo xử lý gấp)", overdue, f"{(overdue/total_tb*100) if total_tb else 0:.1f}%", "Cần báo cáo Lãnh đạo Viện ngay"),
    ]

    ws2.row_dimensions[3].height = 24
    h_kpi = ["STT", "Chỉ Tiêu Đánh Giá Tiến Độ Giải Quyết", "Số Lượng", "Tỷ Lệ %", "Ghi Chú Đôn Đốc Nghiệp Vụ"]
    for i, h in enumerate(h_kpi, start=1):
        c = ws2.cell(row=3, column=i)
        c.value = h
        c.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")

    ws2.column_dimensions["A"].width = 6
    ws2.column_dimensions["B"].width = 50
    ws2.column_dimensions["C"].width = 14
    ws2.column_dimensions["D"].width = 12
    ws2.column_dimensions["E"].width = 40

    for idx, (stt, title, val, pct, note) in enumerate(kpi_rows, start=4):
        ws2.row_dimensions[idx].height = 22
        ws2.cell(row=idx, column=1, value=stt).alignment = Alignment(horizontal="center", vertical="center")
        ws2.cell(row=idx, column=2, value=title).alignment = Alignment(horizontal="left", vertical="center")
        c_val = ws2.cell(row=idx, column=3, value=val)
        c_val.alignment = Alignment(horizontal="center", vertical="center")
        c_val.font = font_data_bold
        ws2.cell(row=idx, column=4, value=pct).alignment = Alignment(horizontal="center", vertical="center")
        ws2.cell(row=idx, column=5, value=note).alignment = Alignment(horizontal="left", vertical="center")

        for c_idx in range(1, 6):
            cell_item = ws2.cell(row=idx, column=c_idx)
            cell_item.border = cell_border
            if stt == "7" and overdue > 0:
                cell_item.fill = fill_overdue
                cell_item.font = Font(name="Segoe UI", size=9, bold=True, color="991B1B")

    # Bảng phân công chi tiết theo KSV
    ksv_stats = {}
    for r in records:
        ksv = r.get("ksv_chinh", "").strip() or "Chưa phân công KSV"
        if ksv not in ksv_stats:
            ksv_stats[ksv] = {"total": 0, "overdue": 0, "urgent": 0, "solved": 0}
        ksv_stats[ksv]["total"] += 1
        if "🔴" in r.get("muc_do_khan_cap", ""):
            ksv_stats[ksv]["overdue"] += 1
        elif "🟠" in r.get("muc_do_khan_cap", ""):
            ksv_stats[ksv]["urgent"] += 1
        elif "✔" in r.get("muc_do_khan_cap", ""):
            ksv_stats[ksv]["solved"] += 1

    sorted_ksv = sorted(ksv_stats.keys(), key=lambda x: ksv_stats[x]["total"], reverse=True)

    r_ksv_start = 13
    ws2.cell(row=r_ksv_start, column=1, value="BẢNG THỐNG KÊ CHI TIẾT KHỐI LƯỢNG TIN BÁO THEO TỪNG KIỂM SÁT VIÊN").font = Font(name="Segoe UI", size=11, bold=True, color="1E3A8A")
    ws2.row_dimensions[r_ksv_start + 1].height = 24

    h_ksv = ["STT", "Họ và Tên Kiểm Sát Viên", "Tổng Số Tin Thụ Lý", "Quá Hạn (🔴)", "Khẩn Cấp (🟠)", "Đã Giải Quyết (✔)", "Tỷ Lệ %"]
    for i, h in enumerate(h_ksv, start=1):
        c = ws2.cell(row=r_ksv_start + 1, column=i)
        c.value = h
        c.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")

    ws2.column_dimensions["F"].width = 18
    ws2.column_dimensions["G"].width = 12

    for k_idx, ksv_name in enumerate(sorted_ksv, start=1):
        curr_row = r_ksv_start + 1 + k_idx
        ws2.row_dimensions[curr_row].height = 20
        st = ksv_stats[ksv_name]
        pct = (st["total"] / total_tb * 100) if total_tb else 0
        vals = [k_idx, ksv_name, st["total"], st["overdue"], st["urgent"], st["solved"], f"{pct:.1f}%"]
        for c_idx, val in enumerate(vals, start=1):
            cell_item = ws2.cell(row=curr_row, column=c_idx, value=val)
            cell_item.border = cell_border
            align = "left" if c_idx == 2 else "center"
            cell_item.alignment = Alignment(horizontal=align, vertical="center")
            if c_idx == 4 and st["overdue"] > 0:
                cell_item.font = Font(name="Segoe UI", size=9, bold=True, color="991B1B")

    # Sheet 3: Ma Trận Pivot Table VKS Thụ Lý × Kết Quả Giải Quyết
    ws3 = wb_out.create_sheet(title="Ma_Tran_Pivot_Tin_Bao")
    ws3.views.sheetView[0].showGridLines = True

    matrix, outcomes, sorted_units = build_pivot_tin_bao_summary(records)

    ws3.merge_cells("A1:J1")
    t3 = ws3["A1"]
    t3.value = "MA TRẬN PIVOT TABLE: ĐƠN VỊ VKS THỤ LÝ × KẾT QUẢ GIẢI QUYẾT TIN BÁO"
    t3.font = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
    t3.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    t3.alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 28

    pivot_headers = ["STT", "Đơn vị VKS thụ lý", "Tổng số tin"] + outcomes + ["Tỷ lệ %"]
    ws3.row_dimensions[3].height = 24
    for i, h in enumerate(pivot_headers, start=1):
        c = ws3.cell(row=3, column=i)
        c.value = h
        c.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")

    ws3.column_dimensions["A"].width = 6
    ws3.column_dimensions["B"].width = 32
    ws3.column_dimensions["C"].width = 14
    for i in range(4, 4 + len(outcomes)):
        ws3.column_dimensions[get_column_letter(i)].width = 16
    ws3.column_dimensions[get_column_letter(4 + len(outcomes))].width = 12

    for u_idx, u in enumerate(sorted_units, start=1):
        curr_row = 3 + u_idx
        ws3.row_dimensions[curr_row].height = 20
        tot = matrix[u]["_total"]
        pct = (tot / total_tb * 100) if total_tb else 0
        vals = [u_idx, u, tot] + [matrix[u].get(k, 0) for k in outcomes] + [f"{pct:.1f}%"]
        for c_idx, val in enumerate(vals, start=1):
            cell_item = ws3.cell(row=curr_row, column=c_idx, value=val)
            cell_item.border = cell_border
            align = "left" if c_idx == 2 else "center"
            cell_item.alignment = Alignment(horizontal=align, vertical="center")

    tot_row_idx = 4 + len(sorted_units)
    ws3.row_dimensions[tot_row_idx].height = 24
    tot_vals = ["∑", "TỔNG CỘNG CHUNG", total_tb] + [sum(matrix[u].get(k, 0) for u in sorted_units) for k in outcomes] + ["100.0%"]
    for c_idx, val in enumerate(tot_vals, start=1):
        cell_item = ws3.cell(row=tot_row_idx, column=c_idx, value=val)
        cell_item.border = cell_border
        cell_item.font = font_data_bold
        cell_item.fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
        align = "left" if c_idx == 2 else "center"
        cell_item.alignment = Alignment(horizontal=align, vertical="center")

    # Sheet 4: Ma Trận Rủi Ro Thời Hiệu TNHS & Hạn Luật Định (Đ.230 BLTTHS)
    ws4 = wb_out.create_sheet(title="Ma_Tran_Rui_Ro_Thoi_Hieu")
    ws4.views.sheetView[0].showGridLines = True

    u_stats_tb, _, sorted_units_tb, _ = build_executive_tin_bao_pivot_summary(records)

    ws4.merge_cells("A1:J1")
    t4 = ws4["A1"]
    t4.value = "MA TRẬN GIÁM SÁT RỦI RO THỜI HIỆU TRUY CỨU TNHS & TIẾN ĐỘ TIN BÁO (THEO Đ.230 BLTTHS)"
    t4.font = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
    t4.fill = PatternFill(start_color="7F1D1D", end_color="7F1D1D", fill_type="solid")
    t4.alignment = Alignment(horizontal="center", vertical="center")
    ws4.row_dimensions[1].height = 28

    risk_headers = [
        ("STT", 6), ("Đơn vị VKS thụ lý", 32), ("Tổng số tin", 14),
        ("Hết thời hiệu ⚠️ (Đ.230)", 24), ("🔴 Khẩn cấp (≤7 ngày)", 20), ("🟠 Rất gấp (8-15 ngày)", 20),
        ("🟡 Sắp hết (16-30 ngày)", 20), ("🟢 An toàn (>30 ngày)", 20), ("⏸️ Tạm đình chỉ", 18), ("Tỷ lệ %", 12)
    ]
    ws4.row_dimensions[3].height = 24
    for i, (h, w) in enumerate(risk_headers, start=1):
        c = ws4.cell(row=3, column=i, value=h)
        c.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color="991B1B", end_color="991B1B", fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws4.column_dimensions[get_column_letter(i)].width = w

    for u_idx, u in enumerate(sorted_units_tb, start=1):
        curr_row = 3 + u_idx
        ws4.row_dimensions[curr_row].height = 20
        st = u_stats_tb[u]
        pct = (st["total_tb"] / total_tb * 100) if total_tb else 0
        vals = [
            u_idx, u, st["total_tb"],
            st["qua_han"] or "", st["khan_cap"] or "", st["rat_gap"] or "",
            st["sap_het"] or "", st["binh_thuong"] or "", st["tam_dinh_chi"] or "",
            f"{pct:.1f}%"
        ]
        for c_idx, val in enumerate(vals, start=1):
            cell_item = ws4.cell(row=curr_row, column=c_idx, value=val)
            cell_item.border = cell_border
            align = "left" if c_idx == 2 else "center"
            cell_item.alignment = Alignment(horizontal=align, vertical="center")
            if c_idx == 4 and st["qua_han"] > 0:
                cell_item.fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
                cell_item.font = Font(name="Segoe UI", size=9, bold=True, color="991B1B")

    tot_r4_idx = 4 + len(sorted_units_tb)
    ws4.row_dimensions[tot_r4_idx].height = 24
    tot_r4_vals = [
        "∑", "TỔNG CỘNG TOÀN VIỆN", total_tb,
        sum(u_stats_tb[u]["qua_han"] for u in sorted_units_tb),
        sum(u_stats_tb[u]["khan_cap"] for u in sorted_units_tb),
        sum(u_stats_tb[u]["rat_gap"] for u in sorted_units_tb),
        sum(u_stats_tb[u]["sap_het"] for u in sorted_units_tb),
        sum(u_stats_tb[u]["binh_thuong"] for u in sorted_units_tb),
        sum(u_stats_tb[u]["tam_dinh_chi"] for u in sorted_units_tb),
        "100.0%"
    ]
    for c_idx, val in enumerate(tot_r4_vals, start=1):
        cell_item = ws4.cell(row=tot_r4_idx, column=c_idx, value=val)
        cell_item.border = cell_border
        cell_item.font = font_data_bold
        cell_item.fill = PatternFill(start_color="FCA5A5", end_color="FCA5A5", fill_type="solid")
        cell_item.alignment = Alignment(horizontal="left" if c_idx == 2 else "center", vertical="center")

    # Bổ sung Biểu đồ Excel trực quan (Data Charts)
    if BarChart and Reference:
        try:
            # 1. Biểu đồ cột tiến độ giải quyết & thời hiệu ở ws2
            chart_td = BarChart()
            chart_td.type = "col"
            chart_td.style = 10
            chart_td.title = "BIỂU ĐỒ TIẾN ĐỘ & MỨC ĐỘ KHẨN CẤP THỜI HIỆU TIN BÁO"
            chart_td.y_axis.title = "Số lượng tin báo"
            chart_td.x_axis.title = "Chỉ tiêu tiến độ"
            data_ref = Reference(ws2, min_col=3, min_row=4, max_row=11)
            cats_ref = Reference(ws2, min_col=2, min_row=5, max_row=11)
            chart_td.add_data(data_ref, titles_from_data=True)
            chart_td.set_categories(cats_ref)
            chart_td.width = 18
            chart_td.height = 10
            ws2.add_chart(chart_td, "I4")

            # 2. Biểu đồ số lượng tin báo theo đơn vị VKS ở ws3
            if sorted_units:
                chart_u = BarChart()
                chart_u.type = "col"
                chart_u.style = 11
                chart_u.title = "BIỂU ĐỒ SỐ LƯỢNG TIN BÁO THEO ĐƠN VỊ VKS THỤ LÝ"
                chart_u.y_axis.title = "Số lượng tin"
                chart_u.x_axis.title = "Đơn vị VKS thụ lý"
                data_u = Reference(ws3, min_col=3, min_row=3, max_row=3 + len(sorted_units))
                cats_u = Reference(ws3, min_col=2, min_row=4, max_row=3 + len(sorted_units))
                chart_u.add_data(data_u, titles_from_data=True)
                chart_u.set_categories(cats_u)
                chart_u.width = 18
                chart_u.height = 11
                col_chart = get_column_letter(5 + len(outcomes))
                ws3.add_chart(chart_u, f"{col_chart}3")
        except Exception:
            pass

    # Sheet Pivot tổng hợp Thời hiệu truy cứu TNHS (Đơn vị VKS x Cảnh báo chi tiết)
    create_tin_bao_thoi_hieu_pivot_sheet(wb_out, records, title="Pivot_Thoi_Hieu_Tin_Bao", index=1)

    wb_out.save(out_filepath)
    try_add_native_excel_pivot_tin_bao(out_filepath, sheet_name="Pivot_Thoi_Hieu_Tin_Bao")
    return out_filepath

def parse_qlahs_case_rows(raw_rows, progress_cb=None, filename=""):
    """Bóc tách dữ liệu Sổ vụ án / Sổ thụ lý / Án tạm đình chỉ thành 24 cột chuẩn."""
    file_context_unit = detect_file_unit(filename, raw_rows) if filename else None
    blocks = []
    current = []
    current_col_map = detect_initial_case_col_map(raw_rows)

    for row in raw_rows:
        if not row:
            continue
        line_str = " ".join(str(c or "").strip().lower() for c in row)
        c0 = row[0] if len(row) > 0 else None
        c1 = row[1] if len(row) > 1 else None

        c0_s = str(c0 or "").strip()
        c1_s = str(c1 or "").strip()

        # Dòng tiêu đề cột (có thể lặp lại nhiều lần khi người dùng copy dán nhiều trang nối tiếp nhau)
        if "thông tin vụ án" in line_str or "mã vụ án" in line_str:
            if current:
                blocks.append((current, current_col_map))
                current = []
            # Cập nhật ngay bản đồ cấu trúc cột tương ứng của bảng mới
            current_col_map = parse_case_header_to_col_map(row)
            continue

        if c0_s.lower() == "stt":
            continue

        # Kiểm tra xem có phải dòng bắt đầu của một vụ án mới (STT dạng số nguyên hoặc số thập phân)
        if re.match(r'^\d+(\.0*)?$', c0_s):
            if current:
                blocks.append((current, current_col_map))
            current = []

        if any(v is not None and str(v).strip() for v in row):
            current.append(row)

    if current:
        blocks.append((current, current_col_map))

    total_blocks = len(blocks)
    records = []

    for idx, (blk, col_map) in enumerate(blocks, 1):
        if progress_cb and total_blocks > 100 and idx % 50 == 0:
            progress_cb(
                f"Đang chuẩn hóa sổ vụ án: {idx:,}/{total_blocks:,} vụ án...",
                "Đang bóc tách số QĐ, thời hiệu, thụ lý và KSV..."
            )

        rec = {k: "" for k, _, _ in COLUMNS_CASE_24}
        rec["stt"] = idx
        ten_lines = []

        for r_idx, row in enumerate(blk):
            def get_cell_val(c_idx):
                if c_idx is not None and len(row) > c_idx and row[c_idx] is not None:
                    return str(row[c_idx]).strip().replace("\xa0", " ")
                return ""

            v_va = get_cell_val(col_map.get("va"))
            v_tl = get_cell_val(col_map.get("tl"))
            v_bc = get_cell_val(col_map.get("bc"))
            v_td = get_cell_val(col_map.get("td"))
            v_lt = get_cell_val(col_map.get("lt"))
            v_cb = get_cell_val(col_map.get("cb"))
            v_tt = get_cell_val(col_map.get("tt"))
            v_bm = get_cell_val(col_map.get("bm"))
            v_dv_moi = get_cell_val(col_map.get("dv_moi"))

            # Lưới an toàn tự chữa lành nếu dữ liệu từng dòng bị xô lệch
            if is_procedural_status_text(v_cb):
                if is_form_status_text(v_tt):
                    v_bm = v_tt
                v_tt = v_cb
                v_cb = ""
            if is_procedural_status_text(v_bm) and not is_procedural_status_text(v_tt) and v_tt:
                v_cb = v_tt
                v_tt = v_bm
                v_bm = ""
            if is_form_status_text(v_tt) and not v_bm:
                v_bm = v_tt
                v_tt = ""

            if v_va:
                va_lines = [ln.strip() for ln in v_va.replace("\r\n", "\n").split("\n") if ln.strip()]
                for l_i, ln in enumerate(va_lines):
                    if r_idx == 0 and l_i == 0 and ln.startswith("VA_"):
                        rec["ma_vu_an"] = ln
                    else:
                        ln_low = ln.lower()
                        if ln_low.startswith("số qđ khởi tố:") or ln_low.startswith("mã số qđ khởi tố"):
                            rec["so_qd_khoi_to"] = ln.split(":", 1)[1].strip()
                        elif ln_low.startswith("ngày ra qđ khởi tố:"):
                            rec["ngay_qd_khoi_to"] = ln.split(":", 1)[1].strip()
                        elif ln_low.startswith("đơn vị ra qđ khởi tố:"):
                            rec["don_vi_khoi_to"] = ln.split(":", 1)[1].strip()
                        elif ln_low.startswith("ngày tiếp nhận hồ sơ:"):
                            rec["ngay_tiep_nhan"] = ln.split(":", 1)[1].strip()
                        elif ln_low.startswith("thời hạn giải quyết:") or ln_low.startswith("thời hạn xử lý:"):
                            rec["thoi_han_giai_quyet"] = ln.split(":", 1)[1].strip()
                        elif ln_low.startswith("ngày hết thời hiệu"):
                            rec["ngay_het_thoi_hieu"] = ln.split(":", 1)[1].strip()
                        elif "thời hiệu" in ln_low or "quá hạn" in ln_low or "sắp hết" in ln_low:
                            rec["canh_bao_thoi_hieu"] = ln
                        elif ln.startswith("VA_") and not rec["ma_vu_an"]:
                            rec["ma_vu_an"] = ln
                        else:
                            ten_lines.append(ln)

            if v_tl:
                tl_lines = [ln.strip() for ln in v_tl.replace("\r\n", "\n").split("\n") if ln.strip()]
                for ln in tl_lines:
                    v_tl_low = ln.lower()
                    if v_tl_low.startswith("số thụ lý:"):
                        rec["so_thu_ly"] = ln.split(":", 1)[1].strip()
                    elif v_tl_low.startswith("ngày thụ lý:"):
                        rec["ngay_thu_ly"] = ln.split(":", 1)[1].strip()
                    elif v_tl_low.startswith("loại thụ lý:"):
                        rec["loai_thu_ly"] = ln.split(":", 1)[1].strip()
                    elif v_tl_low.startswith("đơn vị vks thụ lý:"):
                        rec["vks_thu_ly"] = clean_vks_unit_name(ln.split(":", 1)[1].strip())
                    elif not rec["loai_thu_ly"] and any(x in v_tl_low for x in ("thụ lý mới", "phục hồi", "tiếp nhận")):
                        rec["loai_thu_ly"] = ln

            if v_cb:
                cb_lines = [ln.strip() for ln in v_cb.replace("\r\n", "\n").split("\n") if ln.strip()]
                for ln in cb_lines:
                    v_cb_low = ln.lower()
                    if v_cb_low.startswith("cán bộ thụ lý:"):
                        rec["can_bo_thu_ly"] = ln.split(":", 1)[1].strip()
                    elif v_cb_low.startswith("cán bộ giải quyết:"):
                        rec["can_bo_giai_quyet"] = ln.split(":", 1)[1].strip()
                    else:
                        if rec["can_bo_giai_quyet"]:
                            rec["can_bo_giai_quyet"] += "; " + ln
                        else:
                            rec["can_bo_giai_quyet"] = ln

            if v_bc:
                bc_lines = [ln.strip() for ln in v_bc.replace("\r\n", "\n").split("\n") if ln.strip()]
                for ln in bc_lines:
                    v_bc_low = ln.lower()
                    if v_bc_low.startswith("bị can đầu vụ:"):
                        rec["bi_can_dau_vu"] = ln.split(":", 1)[1].strip()
                    elif v_bc_low.startswith("tổng bị can khởi tố:") or v_bc_low.startswith("tổng số bị can:"):
                        rec["tong_bi_can"] = ln.split(":", 1)[1].strip()
                    elif not rec["bi_can_dau_vu"]:
                        rec["bi_can_dau_vu"] = ln

            if v_td:
                if rec["dieu_luat"]:
                    rec["dieu_luat"] += " | " + v_td
                else:
                    rec["dieu_luat"] = v_td

            if v_lt and not rec["loai_toi"]:
                rec["loai_toi"] = v_lt

            if v_tt:
                rec["trang_thai_vu_an"] = v_tt
                rec["giai_doan"] = detect_procedural_stage(v_tt, filename=filename, loai_thu_ly=rec.get("loai_thu_ly"))

            if v_bm:
                if rec["trang_thai_bieu_mau"]:
                    rec["trang_thai_bieu_mau"] += "; " + v_bm
                else:
                    rec["trang_thai_bieu_mau"] = v_bm

        # Phân tách chính xác Lãnh đạo Viện, Lãnh đạo Phòng, KSV và Cán bộ thụ lý
        ld_v, ld_p, ksv, cb_tl = split_officers(rec.get("can_bo_giai_quyet", ""), rec.get("can_bo_thu_ly", ""))
        rec["lanh_dao_vien"] = ld_v
        rec["lanh_dao_phong"] = ld_p
        rec["can_bo_giai_quyet"] = ksv
        rec["can_bo_thu_ly"] = cb_tl

        rec["ten_vu_an"] = " - ".join(ten_lines)
        prov = detect_province(rec["don_vi_khoi_to"], rec["vks_thu_ly"] + " " + rec["ten_vu_an"])
        if prov in ("Khác / Chưa rõ", "Chưa xác định") and (
            "hồ chí minh" in (rec["vks_thu_ly"] + " " + rec["ten_vu_an"]).lower()
            or "hải quan" in rec["don_vi_khoi_to"].lower()
        ):
            prov = "TP. Hồ Chí Minh"
        rec["tinh_thanh"] = prov

        # Ưu tiên kế thừa cột Tên đơn vị cập nhật mới nhất nếu có
        if v_dv_moi:
            rec["ten_don_vi_moi_nhat"] = v_dv_moi
            u_clean = clean_vks_unit_name(v_dv_moi)
            if u_clean and "chưa" not in u_clean.lower():
                rec["vks_thu_ly"] = u_clean
        elif rec.get("vks_thu_ly"):
            rec["vks_thu_ly"] = clean_vks_unit_name(rec["vks_thu_ly"])

        if not rec.get("vks_thu_ly") or "chưa" in str(rec["vks_thu_ly"]).lower():
            _, u_code = parse_case_code_unit(rec.get("ma_vu_an"), filename)
            if u_code and "chưa" not in str(u_code).lower():
                rec["vks_thu_ly"] = u_code

        # Tính toán mức độ khẩn cấp & số ngày còn lại đến thời hiệu
        update_case_urgency_and_statute(rec, filename=filename)

        # Gán nguồn dữ liệu đối chiếu và xử lý đơn vị từ tên file nếu là mã cũ
        if filename:
            rec["nguon_du_lieu"] = filename
            curr_u = str(rec.get("vks_thu_ly") or "").strip()
            is_unspecific = (
                not curr_u or 
                "chưa xác định" in curr_u.lower() or 
                "chưa rõ" in curr_u.lower() or 
                "mã cũ" in curr_u.lower() or
                curr_u.lower() in ("tp. hồ chí minh", "vksnd thành phố hồ chí minh", "vksnd tp.hcm")
            )
            if file_context_unit and is_unspecific:
                rec["vks_thu_ly"] = file_context_unit

        normalize_case_stage_and_status(rec, filename=filename)
        records.append(rec)

    return records


def select_best_data_sheet(wb):
    """Chọn sheet chứa dữ liệu vụ án / tin báo thực tế thay vì sheet pivot hay tổng quan."""
    sheetnames = wb.sheetnames
    if len(sheetnames) == 1:
        return wb.active

    # 1. Ưu tiên sheet có tên đặc thù chứa dữ liệu danh sách chi tiết
    PRIORITY_NAMES = [
        "danh_sach_hop_nhat_chuan_hoa", "danh_sach_chuan_hoa", "danh_sach_vu_an", 
        "danh_sach_tin_bao", "danh sách", "danh_sach", "data", "sheet1"
    ]
    for p_name in PRIORITY_NAMES:
        for sname in sheetnames:
            if p_name in sname.lower():
                ws = wb[sname]
                if (ws.max_row or 0) > 5:
                    return ws

    # 2. Bỏ qua các sheet Pivot/Báo cáo nếu có sheet khác nhiều dòng hơn
    best_sheet = wb.active
    max_score = -999999
    for sname in sheetnames:
        sname_low = sname.lower()
        ws = wb[sname]
        m_r = ws.max_row or 0
        is_pivot = any(k in sname_low for k in ["pivot", "tong_quan", "tong_hop", "bao_cao", "matrix"])
        score = m_r - (50000 if is_pivot else 0)
        if score > max_score:
            max_score = score
            best_sheet = ws

    return best_sheet


def parse_standardized_case_rows(raw_rows, progress_cb=None, filename=""):
    """
    Bóc tách bảng dữ liệu vụ án phẳng đã chuẩn hóa (26-30 cột do phần mềm xuất ra hoặc người dùng tổng hợp).
    Tự động nhận diện bản đồ cột theo tiêu đề và KẾ THỪA 100% cột 'TÊN ĐƠN VỊ CẬP NHẬT MỚI NHẤT'.
    """
    header_idx = None
    col_map = {}

    for r_i, row in enumerate(raw_rows[:15]):
        if not row:
            continue
        row_str = " ".join(str(c or "").lower() for c in row)
        if "mã vụ án" in row_str or "mã va" in row_str:
            header_idx = r_i
            for c_i, cell in enumerate(row):
                if cell is None:
                    continue
                c_low = str(cell).strip().lower()
                if c_low == "stt":
                    col_map["stt"] = c_i
                elif "nguồn dữ liệu" in c_low or "nguồn đối chiếu" in c_low:
                    col_map["nguon_du_lieu"] = c_i
                elif "mã vụ án" in c_low or "mã va" in c_low:
                    col_map["ma_vu_an"] = c_i
                elif "tên vụ án" in c_low or "nội dung" in c_low:
                    col_map["ten_vu_an"] = c_i
                elif "tỉnh" in c_low or "thành phố" in c_low:
                    col_map["tinh_thanh"] = c_i
                elif "số qđ khởi tố" in c_low or "số qđ" in c_low:
                    col_map["so_qd_khoi_to"] = c_i
                elif "ngày ra qđ" in c_low or "ngày qđ" in c_low:
                    col_map["ngay_qd_khoi_to"] = c_i
                elif "đơn vị ra qđ" in c_low or "cơ quan khởi tố" in c_low:
                    col_map["don_vi_khoi_to"] = c_i
                elif "số thụ lý" in c_low:
                    col_map["so_thu_ly"] = c_i
                elif "ngày thụ lý" in c_low:
                    col_map["ngay_thu_ly"] = c_i
                elif "loại thụ lý" in c_low:
                    col_map["loai_thu_ly"] = c_i
                elif "đơn vị vks thụ lý" in c_low or "vks thụ lý" in c_low or "đơn vị thụ lý" in c_low or "đơn vị vks" in c_low or c_low in ("đơn vị", "vks", "cơ quan thụ lý"):
                    col_map["vks_thu_ly"] = c_i
                elif "ngày tiếp nhận" in c_low:
                    col_map["ngay_tiep_nhan"] = c_i
                elif "thời hạn giải quyết" in c_low or "thời hạn xử lý" in c_low:
                    col_map["thoi_han_giai_quyet"] = c_i
                elif "ngày hết thời hiệu" in c_low:
                    col_map["ngay_het_thoi_hieu"] = c_i
                elif "số ngày còn lại" in c_low:
                    col_map["so_ngay_con_lai"] = c_i
                elif "cảnh báo tiến độ" in c_low or "mức độ khẩn cấp" in c_low or "khẩn cấp" in c_low:
                    col_map["muc_do_khan_cap"] = c_i
                elif "cảnh báo chi tiết" in c_low or "cảnh báo thời hiệu" in c_low:
                    col_map["canh_bao_thoi_hieu"] = c_i
                elif "bị can đầu vụ" in c_low:
                    col_map["bi_can_dau_vu"] = c_i
                elif "tổng số bị can" in c_low or "tổng bị can" in c_low:
                    col_map["tong_bi_can"] = c_i
                elif "tội danh" in c_low or "điều luật" in c_low:
                    col_map["dieu_luat"] = c_i
                elif "phân loại tội" in c_low or "loại tội" in c_low:
                    col_map["loai_toi"] = c_i
                elif "lãnh đạo viện" in c_low:
                    col_map["lanh_dao_vien"] = c_i
                elif "lãnh đạo phòng" in c_low:
                    col_map["lanh_dao_phong"] = c_i
                elif "kiểm sát viên" in c_low or "ksv" in c_low:
                    col_map["can_bo_giai_quyet"] = c_i
                elif "cán bộ thụ lý" in c_low or "chuyên viên" in c_low:
                    col_map["can_bo_thu_ly"] = c_i
                elif "giai đoạn" in c_low:
                    col_map["giai_doan"] = c_i
                elif "trạng thái vụ án" in c_low:
                    col_map["trang_thai_vu_an"] = c_i
                elif "trạng thái biểu mẫu" in c_low:
                    col_map["trang_thai_bieu_mau"] = c_i
                elif "mới nhất" in c_low or "cập nhật mới" in c_low or "đơn vị cập nhật" in c_low:
                    col_map["ten_don_vi_moi_nhat"] = c_i
            break

    if header_idx is None:
        return []

    records = []
    for r_i, row in enumerate(raw_rows[header_idx + 1:], 1):
        if not row or not any(row):
            continue

        def get_val(key):
            idx = col_map.get(key)
            if idx is not None and len(row) > idx and row[idx] is not None:
                return str(row[idx]).strip().replace("\xa0", " ")
            return ""

        ma_va = get_val("ma_vu_an")
        if not ma_va or not (ma_va.startswith("VA_") or "VA" in ma_va):
            continue

        rec = {k: "" for k, _, _ in COLUMNS_CASE_29}
        for k in rec.keys():
            rec[k] = get_val(k)

        # Định dạng ngày chuẩn DD/MM/YYYY
        for dk in ("ngay_qd_khoi_to", "ngay_thu_ly", "ngay_tiep_nhan", "thoi_han_giai_quyet", "ngay_het_thoi_hieu"):
            if rec.get(dk):
                rec[dk] = format_date_str(rec[dk])

        # Kế thừa nguồn dữ liệu
        rec["nguon_du_lieu"] = get_val("nguon_du_lieu") or filename or "Tập tin đã chuẩn hóa"

        # KẾ THỪA CỘT TÊN ĐƠN VỊ CẬP NHẬT MỚI NHẤT (CHO VỤ ÁN)
        don_vi_moi = get_val("ten_don_vi_moi_nhat")
        if don_vi_moi:
            rec["ten_don_vi_moi_nhat"] = don_vi_moi
            u_std = clean_vks_unit_name(don_vi_moi)
            if u_std and "chưa" not in u_std.lower():
                rec["vks_thu_ly"] = u_std
        elif rec.get("vks_thu_ly"):
            rec["vks_thu_ly"] = clean_vks_unit_name(rec["vks_thu_ly"])

        if not rec.get("vks_thu_ly") or "chưa" in str(rec["vks_thu_ly"]).lower():
            _, u_code = parse_case_code_unit(rec.get("ma_vu_an"), filename)
            if u_code and "chưa" not in str(u_code).lower():
                rec["vks_thu_ly"] = u_code

        normalize_case_stage_and_status(rec, filename=filename)
        records.append(rec)

    return records


def parse_standardized_tin_bao_rows(raw_rows, progress_cb=None, filename=""):
    """
    Bóc tách bảng dữ liệu tin báo phẳng đã chuẩn hóa (35-39 cột do phần mềm xuất ra hoặc người dùng tổng hợp).
    Tự động nhận diện bản đồ cột theo tiêu đề, kế thừa 100% cột 'TÊN ĐƠN VỊ CẬP NHẬT MỚI NHẤT',
    và tính toán đồng bộ cảnh báo thời hiệu truy cứu TNHS theo đúng quy định.
    """
    header_idx = None
    col_map = {}

    for r_i, row in enumerate(raw_rows[:15]):
        if not row:
            continue
        row_str = " ".join(str(c or "").lower() for c in row)
        has_exact_col = any(str(c or "").strip().lower() in ("mã tin báo", "mã tb", "mã tin") for c in row)
        if ("mã tin báo" in row_str or "mã tb" in row_str or "mã tin" in row_str) and has_exact_col:
            header_idx = r_i
            for c_i, cell in enumerate(row):
                if cell is None:
                    continue
                c_low = str(cell).strip().lower()
                if c_low == "stt":
                    col_map["stt"] = c_i
                elif "nguồn dữ liệu" in c_low or "nguồn đối chiếu" in c_low:
                    col_map["nguon_du_lieu"] = c_i
                elif "mã tin báo" in c_low or "mã tb" in c_low or "mã tin" in c_low:
                    col_map["ma_tin_bao"] = c_i
                elif "loại tin" in c_low:
                    col_map["loai_tin"] = c_i
                elif "phân loại nguồn" in c_low or "nguồn tin" in c_low or "phân loại" in c_low:
                    col_map["nguon_tin"] = c_i
                elif "phức tạp" in c_low:
                    col_map["tinh_tiet_phuc_tap"] = c_i
                elif "người báo tin" in c_low or "cơ quan báo tin" in c_low or "cq báo tin" in c_low:
                    col_map["nguoi_bao_tin"] = c_i
                elif "người bị tố giác" in c_low or "đối tượng bị tố giác" in c_low or "bị tố giác" in c_low:
                    col_map["nguoi_bi_to_giac"] = c_i
                elif "cccd" in c_low or "định danh" in c_low:
                    col_map["cccd_nguoi_bao"] = c_i
                elif "số điện thoại" in c_low or "sđt" in c_low:
                    col_map["sdt_nguoi_bao"] = c_i
                elif "địa chỉ người báo" in c_low or "địa chỉ" in c_low:
                    col_map["dia_chi_nguoi_bao"] = c_i
                elif "thời gian & địa điểm" in c_low or "thời gian" in c_low:
                    col_map["thoi_gian_dia_diem"] = c_i
                elif "nội dung" in c_low:
                    col_map["noi_dung_tin_bao"] = c_i
                elif "điều luật" in c_low or "tội danh" in c_low or c_low == "tội":
                    col_map["dieu_luat"] = c_i
                elif "chương blhs" in c_low or "chương" in c_low:
                    col_map["chuong_toi"] = c_i
                elif "mức độ tội" in c_low or "mức độ nghiêm trọng" in c_low:
                    col_map["muc_do_nghiem_trong"] = c_i
                elif "đơn vị tiếp nhận" in c_low:
                    col_map["don_vi_tiep_nhan"] = c_i
                elif "ngày tiếp nhận" in c_low or ("tiếp nhận" in c_low and "ngày" in c_low):
                    col_map["ngay_tiep_nhan_tin"] = c_i
                elif "ngày cqđt" in c_low or "cqđt thụ lý" in c_low or "ngày cqđt thụ lý" in c_low:
                    col_map["ngay_cqdt_nhan"] = c_i
                elif "điều tra viên" in c_low or "đtv" in c_low:
                    col_map["dieu_tra_vien"] = c_i
                elif "vks tiếp nhận" in c_low or "vks nhận" in c_low or "ngày vks nhận" in c_low:
                    col_map["ngay_vks_nhan"] = c_i
                elif "số thụ lý" in c_low or "số tl" in c_low:
                    col_map["so_thu_ly"] = c_i
                elif "ngày vks thụ lý" in c_low:
                    col_map["ngay_vks_thu_ly"] = c_i
                elif "đơn vị vks thụ lý" in c_low or "vks thụ lý" in c_low or "đơn vị thụ lý" in c_low or "đơn vị vks" in c_low or c_low in ("đơn vị", "vks", "cơ quan thụ lý"):
                    col_map["vks_thu_ly"] = c_i
                elif "ksv phối hợp" in c_low or "kiểm sát viên phối hợp" in c_low or "phối hợp" in c_low or "giúp việc" in c_low:
                    col_map["ksv_phoi_hop"] = c_i
                elif "ksv thụ lý chính" in c_low or "kiểm sát viên thụ lý" in c_low or "ksv thụ lý" in c_low or "kiểm sát viên chính" in c_low or "ksv chính" in c_low or "cán bộ thụ lý" in c_low or "cán bộ giải quyết" in c_low or "cán bộ kiểm sát" in c_low or "kiểm sát viên" in c_low or c_low == "ksv":
                    if "ksv_chinh" not in col_map:
                        col_map["ksv_chinh"] = c_i
                    else:
                        col_map["ksv_phoi_hop"] = c_i
                elif "đơn vị giải quyết" in c_low:
                    col_map["don_vi_giai_quyet"] = c_i
                elif "thời hạn từ ngày" in c_low or "từ ngày" in c_low:
                    col_map["thoi_han_tu_ngay"] = c_i
                elif "thời hạn đến ngày" in c_low or "đến ngày" in c_low or "hết hạn giải quyết" in c_low or "hạn giải quyết" in c_low:
                    col_map["thoi_han_den_ngay"] = c_i
                elif "số ngày còn lại thời hiệu" in c_low:
                    col_map["so_ngay_con_lai_thoi_hieu"] = c_i
                elif "số ngày còn lại" in c_low and "so_ngay_con_lai" not in col_map:
                    col_map["so_ngay_con_lai"] = c_i
                elif "cảnh báo tiến độ" in c_low or "mức độ khẩn cấp" in c_low or "tiến độ / thời hạn" in c_low:
                    col_map["muc_do_khan_cap"] = c_i
                elif "kết quả giải quyết" in c_low or "tình trạng" in c_low:
                    col_map["ket_qua_giai_quyet"] = c_i
                elif "ngày tạm đình chỉ" in c_low:
                    col_map["ngay_tam_dinh_chi"] = c_i
                elif "lý do tạm đình chỉ" in c_low:
                    col_map["ly_do_tam_dinh_chi"] = c_i
                elif "ngày hết thời hiệu" in c_low:
                    col_map["ngay_het_thoi_hieu"] = c_i
                elif "cảnh báo thời hiệu" in c_low:
                    col_map["canh_bao_thoi_hieu"] = c_i
                elif "cảnh báo chi tiết" in c_low:
                    col_map["canh_bao_chi_tiet_thoi_hieu"] = c_i
                elif "ghi chú" in c_low:
                    col_map["ghi_chu"] = c_i
                elif "mới nhất" in c_low or "cập nhật mới" in c_low or "đơn vị cập nhật" in c_low:
                    col_map["ten_don_vi_moi_nhat"] = c_i
            break

    if header_idx is None:
        return []

    today_date = date.today()
    records = []
    for r_i, row in enumerate(raw_rows[header_idx + 1:], 1):
        if not row or not any(row):
            continue

        def get_val(key):
            idx = col_map.get(key)
            if idx is not None and len(row) > idx and row[idx] is not None:
                return str(row[idx]).strip().replace("\xa0", " ")
            return ""

        ma_tb = get_val("ma_tin_bao")
        so_tl = get_val("so_thu_ly")
        noi_dung = get_val("noi_dung_tin_bao")
        nguoi_bao = get_val("nguoi_bao_tin")

        if not ma_tb and not so_tl and not noi_dung and not nguoi_bao:
            continue
        if ma_tb.lower() in ("mã tin báo", "mã tin", "stt"):
            continue

        rec = {k: "" for k, _, _ in COLUMNS_TIN_BAO_35}
        for k in rec.keys():
            rec[k] = get_val(k)

        # Chuẩn hóa ngày tháng
        for dk in ("ngay_tiep_nhan_tin", "ngay_cqdt_nhan", "ngay_vks_nhan", "ngay_vks_thu_ly",
                   "thoi_han_tu_ngay", "thoi_han_den_ngay", "ngay_tam_dinh_chi", "ngay_het_thoi_hieu"):
            if rec.get(dk):
                rec[dk] = format_date_str(rec[dk])

        # Kế thừa nguồn dữ liệu
        rec["nguon_du_lieu"] = get_val("nguon_du_lieu") or filename or "Tập tin đã chuẩn hóa"

        # Đơn vị VKS thụ lý (Danh sách tin báo đã có sẵn đầy đủ)
        don_vi_moi = get_val("ten_don_vi_moi_nhat")
        if don_vi_moi:
            u_std = clean_vks_unit_name(don_vi_moi)
            if u_std and "chưa" not in u_std.lower():
                rec["vks_thu_ly"] = u_std
        elif rec.get("vks_thu_ly"):
            rec["vks_thu_ly"] = clean_vks_unit_name(rec["vks_thu_ly"])

        if not rec.get("vks_thu_ly") or "chưa" in str(rec["vks_thu_ly"]).lower():
            u_code = parse_tin_bao_code_unit(rec.get("ma_tin_bao"), filename)
            if u_code and "chưa" not in str(u_code).lower():
                rec["vks_thu_ly"] = u_code
        rec.pop("ten_don_vi_moi_nhat", None)

        # Chuẩn hóa tên cán bộ Kiểm sát & Điều tra viên
        raw_ksv = rec.get("ksv_chinh")
        if raw_ksv:
            if not rec.get("ksv_phoi_hop") and any(sep in str(raw_ksv) for sep in (",", ";", "\n", "/")):
                parts = [clean_person_name(p) for p in re.split(r'[,;\n/]', str(raw_ksv)) if clean_person_name(p)]
                if parts:
                    rec["ksv_chinh"] = parts[0]
                    if len(parts) > 1:
                        rec["ksv_phoi_hop"] = "; ".join(parts[1:])
            else:
                rec["ksv_chinh"] = clean_person_name(raw_ksv)
        if rec.get("ksv_phoi_hop"):
            rec["ksv_phoi_hop"] = clean_person_name(rec["ksv_phoi_hop"])
        if rec.get("dieu_tra_vien"):
            rec["dieu_tra_vien"] = clean_person_name(rec["dieu_tra_vien"])

        # Tính toán & đồng bộ cảnh báo thời hiệu / tiến độ
        th_raw = rec.get("ngay_het_thoi_hieu")
        d_th = parse_date_safely(th_raw)
        kq_raw = rec.get("ket_qua_giai_quyet") or "Đang giải quyết"

        if d_th:
            days_th = (d_th - today_date).days
            if days_th < 0:
                rec["so_ngay_con_lai_thoi_hieu"] = f"Đã quá {-days_th} ngày"
                rec["canh_bao_thoi_hieu"] = f"⚠️ Đã quá hạn ({-days_th} ngày)"
                rec["canh_bao_chi_tiet_thoi_hieu"] = "Đã hết thời hiệu truy cứu TNHS"
                rec["so_ngay_con_lai"] = f"Đã quá {-days_th} ngày"
                rec["muc_do_khan_cap"] = f"⚠️ Đã hết thời hiệu truy cứu TNHS (Quá {-days_th} ngày)"
            elif days_th == 0:
                rec["so_ngay_con_lai_thoi_hieu"] = "Hôm nay hết hạn"
                rec["canh_bao_thoi_hieu"] = "🔴 Hết thời hiệu hôm nay (0 ngày)"
                rec["canh_bao_chi_tiet_thoi_hieu"] = "Hết thời hiệu hôm nay"
                rec["so_ngay_con_lai"] = "Hôm nay hết hạn"
                rec["muc_do_khan_cap"] = "🔴 Hết thời hiệu hôm nay (0 ngày)"
            elif days_th <= 7:
                rec["so_ngay_con_lai_thoi_hieu"] = f"Còn {days_th} ngày"
                rec["canh_bao_thoi_hieu"] = f"🔴 Khẩn cấp (Còn {days_th} ngày)"
                rec["canh_bao_chi_tiet_thoi_hieu"] = "Sắp hết thời hiệu truy cứu TNHS"
                rec["so_ngay_con_lai"] = f"Còn {days_th} ngày"
                rec["muc_do_khan_cap"] = f"🔴 Sắp hết thời hiệu truy cứu TNHS (Còn {days_th} ngày)"
            elif days_th <= 15:
                rec["so_ngay_con_lai_thoi_hieu"] = f"Còn {days_th} ngày"
                rec["canh_bao_thoi_hieu"] = f"🟠 Rất gấp (Còn {days_th} ngày)"
                rec["canh_bao_chi_tiet_thoi_hieu"] = "Sắp hết thời hiệu truy cứu TNHS"
                rec["so_ngay_con_lai"] = f"Còn {days_th} ngày"
                rec["muc_do_khan_cap"] = f"🟠 Sắp hết thời hiệu truy cứu TNHS (Còn {days_th} ngày)"
            elif days_th <= 30:
                rec["so_ngay_con_lai_thoi_hieu"] = f"Còn {days_th} ngày"
                rec["canh_bao_thoi_hieu"] = f"🟡 Sắp hết hạn (Còn {days_th} ngày)"
                rec["canh_bao_chi_tiet_thoi_hieu"] = "Sắp hết thời hiệu truy cứu TNHS"
                rec["so_ngay_con_lai"] = f"Còn {days_th} ngày"
                rec["muc_do_khan_cap"] = f"🟡 Sắp hết thời hiệu truy cứu TNHS (Còn {days_th} ngày)"
            else:
                rec["so_ngay_con_lai_thoi_hieu"] = f"Còn {days_th} ngày"
                rec["canh_bao_thoi_hieu"] = f"🟢 Bình thường (Còn {days_th} ngày)"
                rec["canh_bao_chi_tiet_thoi_hieu"] = "Trong thời hiệu truy cứu TNHS"
                if "tạm đình chỉ" in kq_raw.lower() or rec.get("ngay_tam_dinh_chi"):
                    rec["so_ngay_con_lai"] = f"Còn {days_th} ngày"
                    rec["muc_do_khan_cap"] = f"⏸ Tạm đình chỉ (Còn {days_th} ngày thời hiệu)"
        elif "không tính" in str(th_raw).lower():
            rec["so_ngay_con_lai_thoi_hieu"] = "Không tính thời hiệu"
            rec["canh_bao_thoi_hieu"] = "⚪ Không tính thời hiệu"
            rec["canh_bao_chi_tiet_thoi_hieu"] = "Không tính thời hiệu truy cứu"
        elif "tạm đình chỉ" in kq_raw.lower() or rec.get("ngay_tam_dinh_chi"):
            rec["so_ngay_con_lai"] = "Tạm đình chỉ"
            rec["muc_do_khan_cap"] = "⏸ Tạm đình chỉ giải quyết"
            rec["so_ngay_con_lai_thoi_hieu"] = "Chưa có ngày"
            rec["canh_bao_thoi_hieu"] = "⚠️ Tạm đình chỉ (Chưa rõ thời hiệu)"
            rec["canh_bao_chi_tiet_thoi_hieu"] = "Cần rà soát bổ sung ngày hết thời hiệu"

        normalize_tin_bao_status(rec)
        records.append(rec)

    return records


def load_excel_and_parse(filepath, progress_cb=None):
    """Đọc file Excel và chuẩn hóa dữ liệu, hỗ trợ tiến trình và tối ưu luồng bộ nhớ."""
    if progress_cb:
        progress_cb("Đang mở và đọc cấu trúc tập tin Excel...", os.path.basename(filepath))

    raw_rows = []
    # Thử đọc nhanh bằng read_only=True
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        ws = select_best_data_sheet(wb)
        count = 0
        for row in ws.iter_rows(values_only=True):
            count += 1
            raw_rows.append(list(row) if row else [])
            if progress_cb and count % 500 == 0:
                progress_cb(
                    f"Đang tải dữ liệu Excel: đã nạp {count:,} dòng...",
                    os.path.basename(filepath)
                )
        wb.close()
    except Exception:
        # Dự phòng mở tiêu chuẩn nếu file có định dạng đặc thù
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = select_best_data_sheet(wb)
        max_r = ws.max_row
        max_c = min(ws.max_column, 80)
        for r in range(1, max_r + 1):
            raw_rows.append([ws.cell(row=r, column=c).value for c in range(1, max_c + 1)])
            if progress_cb and r % 500 == 0:
                progress_cb(
                    f"Đang tải dữ liệu Excel: {r:,}/{max_r:,} dòng...",
                    os.path.basename(filepath)
                )
        wb.close()

    fn_base = os.path.basename(filepath)
    fmt = detect_data_format(raw_rows)
    if fmt == "STANDARDIZED_CASE":
        if progress_cb:
            progress_cb(
                f"Phát hiện Danh sách Vụ án chuẩn hóa / tổng hợp ({len(raw_rows):,} dòng)...",
                "Đang kế thừa đơn vị cập nhật mới nhất, phân công KSV và tiến độ..."
            )
        recs = parse_standardized_case_rows(raw_rows, progress_cb, filename=fn_base)
        fmt = "CASE"
    elif fmt == "STANDARDIZED_TIN_BAO":
        if progress_cb:
            progress_cb(
                f"Phát hiện Danh sách Tin báo chuẩn hóa / tổng hợp ({len(raw_rows):,} dòng)...",
                "Đang kế thừa đơn vị cập nhật mới nhất và rà soát thời hiệu..."
            )
        recs = parse_standardized_tin_bao_rows(raw_rows, progress_cb, filename=fn_base)
        fmt = "TIN_BAO"
    elif fmt == "TIN_BAO_RAW":
        if progress_cb:
            progress_cb(
                f"Phát hiện định dạng Sổ tin báo copy trực tiếp ({len(raw_rows):,} dòng)...",
                "Đang bóc tách theo khối tin báo, trích xuất KSV & thời hạn..."
            )
        recs = parse_qlahs_raw_tin_bao_copy_rows(raw_rows, progress_cb, filename=fn_base)
        fmt = "TIN_BAO"
    elif fmt == "TIN_BAO":
        if progress_cb:
            progress_cb(
                f"Phát hiện định dạng Danh sách Tin báo về tội phạm ({len(raw_rows):,} dòng)...",
                "Đang bóc tách 30 cột chuẩn hóa, thời hạn & kiểm sát viên..."
            )
        recs = parse_qlahs_tin_bao_rows(raw_rows, progress_cb, filename=fn_base)
    elif fmt == "CASE_REPORT":
        if progress_cb:
            progress_cb(
                f"Phát hiện định dạng Báo cáo Hồ sơ vụ án ({len(raw_rows):,} dòng)...",
                "Đang bóc tách thông tin vụ án, hạn điều tra & cán bộ..."
            )
        recs = parse_qlahs_case_report_rows(raw_rows, progress_cb, filename=fn_base)
    elif fmt == "CASE":
        if progress_cb:
            progress_cb(
                f"Phát hiện định dạng Sổ Vụ án / Thụ lý ({len(raw_rows):,} dòng)...",
                "Đang bóc tách 24 cột chuẩn hóa chi tiết..."
            )
        recs = parse_qlahs_case_rows(raw_rows, progress_cb, filename=fn_base)
    else:
        if progress_cb:
            progress_cb(
                f"Phát hiện định dạng Sổ Bị can ({len(raw_rows):,} dòng)...",
                "Đang bóc tách 12 cột chuẩn hóa..."
            )
        recs = parse_qlahs_rows(raw_rows, progress_cb, filename=fn_base)

    if fmt in ("CASE", "CASE_REPORT", "STANDARDIZED_CASE"):
        for r in recs:
            normalize_case_stage_and_status(r, filename=fn_base)

    return fmt, recs


def parse_clipboard_data(clip_text, progress_cb=None):
    """Phân tích dữ liệu dán từ clipboard và tự động nhận diện định dạng."""
    if not clip_text:
        return "DEFENDANT", []
    lines = [l for l in clip_text.replace("\r\n", "\n").split("\n") if l.strip()]
    if not lines:
        return "DEFENDANT", []

    # Kiểm tra delimiter (tab hay comma)
    sample_first = lines[0]
    import csv, io
    if "\t" in sample_first:
        raw_rows = [l.split("\t") for l in lines]
    else:
        try:
            reader = csv.reader(io.StringIO(clip_text))
            raw_rows = [r for r in reader if r]
        except Exception:
            raw_rows = [l.split(",") for l in lines]

    fmt = detect_data_format(raw_rows)
    if fmt == "STANDARDIZED_CASE":
        if progress_cb:
            progress_cb(f"Nhận diện bảng Vụ án chuẩn hóa ({len(raw_rows):,} dòng)...", "Đang kế thừa đơn vị cập nhật mới nhất...")
        recs = parse_standardized_case_rows(raw_rows, progress_cb, filename="Clipboard")
        fmt = "CASE"
    elif fmt == "STANDARDIZED_TIN_BAO":
        if progress_cb:
            progress_cb(f"Nhận diện bảng Tin báo chuẩn hóa ({len(raw_rows):,} dòng)...", "Đang kế thừa đơn vị cập nhật mới nhất...")
        recs = parse_standardized_tin_bao_rows(raw_rows, progress_cb, filename="Clipboard")
        fmt = "TIN_BAO"
    elif fmt == "TIN_BAO":
        if progress_cb:
            progress_cb(f"Nhận diện định dạng Tin báo tội phạm ({len(raw_rows):,} dòng)...", "Đang bóc tách 35 cột...")
        recs = parse_qlahs_tin_bao_rows(raw_rows, progress_cb)
    elif fmt == "CASE_REPORT":
        if progress_cb:
            progress_cb(f"Nhận diện định dạng Báo cáo Hồ sơ vụ án ({len(raw_rows):,} dòng)...", "Đang bóc tách thông tin vụ án & thời hạn...")
        recs = parse_qlahs_case_report_rows(raw_rows, progress_cb, filename="Clipboard")
    elif fmt == "CASE":
        if progress_cb:
            progress_cb(f"Nhận diện định dạng Sổ Vụ án ({len(raw_rows):,} dòng)...", "Đang bóc tách cột chuẩn...")
        recs = parse_qlahs_case_rows(raw_rows, progress_cb, filename="Clipboard")
    else:
        if progress_cb:
            progress_cb(f"Nhận diện định dạng Sổ Bị can ({len(raw_rows):,} dòng)...", "Đang bóc tách 12 cột...")
        recs = parse_qlahs_rows(raw_rows, progress_cb, filename="Clipboard")

    if fmt in ("CASE", "CASE_REPORT", "STANDARDIZED_CASE"):
        for r in recs:
            normalize_case_stage_and_status(r)

    return fmt, recs


def build_pivot_case_summary(records):
    """Tính toán ma trận Pivot Table cho Sổ vụ án: Đơn vị VKS thụ lý × Phân loại tội phạm."""
    if not records:
        return {}, [], []

    crime_types_order = ["Ít nghiêm trọng", "Nghiêm trọng", "Rất nghiêm trọng", "Đặc biệt nghiêm trọng"]
    found_types = []
    for r in records:
        lt = r.get("loai_toi", "").strip() or "Chưa phân loại"
        if lt not in found_types:
            found_types.append(lt)

    crime_cols = [c for c in crime_types_order if c in found_types]
    for c in found_types:
        if c not in crime_cols:
            crime_cols.append(c)

    matrix = {}
    for r in records:
        unit = r.get("vks_thu_ly", "").strip() or "Chưa rõ đơn vị"
        lt = r.get("loai_toi", "").strip() or "Chưa phân loại"

        if unit not in matrix:
            matrix[unit] = {
                "total_cases": 0,
                "crimes": {c: 0 for c in crime_cols},
            }
        matrix[unit]["total_cases"] += 1
        if lt not in matrix[unit]["crimes"]:
            matrix[unit]["crimes"][lt] = 0
        matrix[unit]["crimes"][lt] += 1

    sorted_units = sorted(matrix.keys(), key=lambda u: matrix[u]["total_cases"], reverse=True)
    return matrix, crime_cols, sorted_units


def build_pivot_stage_summary(records):
    """Tính toán ma trận Pivot Table: Đơn vị VKS thụ lý × Giai đoạn tố tụng (Điều tra, Truy tố, Sơ thẩm, Phúc thẩm...)."""
    if not records:
        return {}, [], []

    stages_order = ["Điều tra", "Truy tố", "Sơ thẩm", "Phúc thẩm", "Thi hành án"]
    found_stages = []
    for r in records:
        gd = r.get("giai_doan", "").strip() or "Điều tra"
        if gd not in found_stages:
            found_stages.append(gd)

    active_stages = [s for s in stages_order if s in found_stages]
    for s in found_stages:
        if s not in active_stages:
            active_stages.append(s)

    if not active_stages:
        active_stages = ["Điều tra", "Truy tố", "Sơ thẩm"]

    matrix = {}
    for r in records:
        u = r.get("vks_thu_ly", "").strip() or "Chưa rõ đơn vị VKS"
        gd = r.get("giai_doan", "").strip() or "Điều tra"
        if u not in matrix:
            matrix[u] = {
                "total_cases": 0,
                "stages": {s: 0 for s in active_stages},
            }
        matrix[u]["total_cases"] += 1
        if gd not in matrix[u]["stages"]:
            matrix[u]["stages"][gd] = 0
        matrix[u]["stages"][gd] += 1

    sorted_units = sorted(matrix.keys(), key=lambda u: matrix[u]["total_cases"], reverse=True)
    return matrix, active_stages, sorted_units


def build_executive_master_pivot_summary(records):
    """Tính toán ma trận đa chiều 360° phục vụ Lãnh đạo Viện / Lãnh đạo Phòng."""
    stages_order = ["Điều tra", "Truy tố", "Sơ thẩm", "Phúc thẩm"]
    found_stages = set()
    for r in records:
        gd = r.get("giai_doan", "").strip() or "Điều tra"
        found_stages.add(gd)

    active_stages = [s for s in stages_order if s in found_stages]
    for s in sorted(found_stages):
        if s not in active_stages and s not in ("Chưa rõ", ""):
            active_stages.append(s)
    if not active_stages:
        active_stages = ["Điều tra", "Truy tố", "Sơ thẩm"]

    unit_stats = {}
    tot_all = len(records)
    for r in records:
        u = r.get("vks_thu_ly", "").strip() or "Chưa rõ đơn vị VKS"
        gd = r.get("giai_doan", "").strip() or "Điều tra"
        urg = (str(r.get("muc_do_khan_cap") or "") + " " + str(r.get("canh_bao_thoi_hieu") or "")).lower()
        st_va = str(r.get("trang_thai_vu_an") or "").lower()
        lt = str(r.get("loai_toi") or "").lower()

        # Số lượng bị can
        tbc_str = str(r.get("tong_bi_can") or "").strip()
        tbc = int(tbc_str) if tbc_str.isdigit() else (1 if r.get("bi_can_dau_vu") else 0)

        if u not in unit_stats:
            unit_stats[u] = {
                "total_cases": 0,
                "total_defendants": 0,
                "stages": {s: 0 for s in active_stages},
                "qua_han": 0,
                "khan_cap": 0,
                "rat_gap": 0,
                "sap_het": 0,
                "binh_thuong": 0,
                "tam_dinh_chi": 0,
                "toi_it": 0,
                "toi_nt": 0,
                "toi_rnt": 0,
                "toi_dbnt": 0,
            }

        st = unit_stats[u]
        st["total_cases"] += 1
        st["total_defendants"] += tbc
        if gd not in st["stages"]:
            st["stages"][gd] = 0
        st["stages"][gd] += 1

        # Phân loại rủi ro thời hiệu & tiến độ (Đ.230 BLTTHS & Đ.27 BLHS)
        # ƯU TIÊN KIỂM TRA CẢNH BÁO THỜI HIỆU TRUY CỨU TNHS TRƯỚC HẾT ĐỂ TRÁNH BỊ TẠM ĐÌNH CHỈ NUỐT MẤT
        is_expired = ("⚠️" in urg or "đã hết" in urg or "đã quá" in urg or "quá hạn" in urg or "quá thời hiệu" in urg or ("hết thời hiệu" in urg and "sắp" not in urg))
        if is_expired:
            st["qua_han"] += 1
        elif "🔴" in urg or "khẩn cấp" in urg:
            st["khan_cap"] += 1
        elif "🟠" in urg or "rất gấp" in urg:
            st["rat_gap"] += 1
        elif "🟡" in urg or "sắp hết" in urg:
            st["sap_het"] += 1
        elif "tạm đình chỉ" in st_va or "tạm đình chỉ" in urg or "⏸" in urg:
            st["tam_dinh_chi"] += 1
        elif any(k in st_va or k in urg for k in ("đình chỉ", "xét xử", "hoàn thành", "đã giải quyết", "kết thúc điều tra", "cáo trạng")):
            pass
        else:
            st["binh_thuong"] += 1

        # Phân loại tội danh
        if "ít" in lt:
            st["toi_it"] += 1
        elif "rất" in lt:
            st["toi_rnt"] += 1
        elif "đặc biệt" in lt:
            st["toi_dbnt"] += 1
        elif "nghiêm trọng" in lt:
            st["toi_nt"] += 1

    sorted_units = sorted(unit_stats.keys(), key=lambda x: unit_stats[x]["total_cases"], reverse=True)
    return unit_stats, active_stages, sorted_units, tot_all


def build_executive_tin_bao_pivot_summary(records):
    """Tính toán ma trận đa chiều phục vụ Lãnh đạo Viện / Lãnh đạo Phòng đối với nguồn Tin báo."""
    common_order = ["Khởi tố", "Không khởi tố", "Tạm đình chỉ", "Đang giải quyết", "Chuyển nơi khác", "Không thụ lý", "Đã nhập"]
    found_results = set()
    for r in records:
        kq = r.get("ket_qua_giai_quyet", "").strip() or "Đang giải quyết"
        found_results.add(kq)

    active_outcomes = [x for x in common_order if x in found_results]
    for x in sorted(found_results):
        if x not in active_outcomes:
            active_outcomes.append(x)
    if not active_outcomes:
        active_outcomes = ["Khởi tố", "Không khởi tố", "Tạm đình chỉ", "Đang giải quyết"]

    unit_stats = {}
    tot_all = len(records)
    for r in records:
        u = r.get("vks_thu_ly", "").strip() or "Chưa rõ đơn vị VKS"
        kq = r.get("ket_qua_giai_quyet", "").strip() or "Đang giải quyết"
        urg = (str(r.get("muc_do_khan_cap") or "") + " " + str(r.get("canh_bao_thoi_hieu") or "")).lower()

        if u not in unit_stats:
            unit_stats[u] = {
                "total_tb": 0,
                "outcomes": {k: 0 for k in active_outcomes},
                "qua_han": 0,      # Đã hết thời hiệu truy cứu TNHS
                "khan_cap": 0,     # Còn <= 7 ngày
                "rat_gap": 0,      # Còn 8-15 ngày
                "sap_het": 0,      # Còn 16-30 ngày
                "binh_thuong": 0,  # Trong thời hiệu / An toàn
                "tam_dinh_chi": 0, # Tạm đình chỉ giải quyết
                "da_giai_quyet": 0 # Khởi tố / Không khởi tố...
            }

        st = unit_stats[u]
        st["total_tb"] += 1
        if kq not in st["outcomes"]:
            st["outcomes"][kq] = 0
        st["outcomes"][kq] += 1

        # Phân loại rủi ro thời hiệu & tiến độ (Đ.230 BLTTHS)
        is_expired = ("⚠️" in urg or "đã hết" in urg or "đã quá" in urg or "quá hạn" in urg or "quá thời hiệu" in urg or ("hết thời hiệu" in urg and "sắp" not in urg))
        if is_expired:
            st["qua_han"] += 1
        elif "🔴" in urg or "khẩn cấp" in urg:
            st["khan_cap"] += 1
        elif "🟠" in urg or "rất gấp" in urg:
            st["rat_gap"] += 1
        elif "🟡" in urg or "sắp hết" in urg:
            st["sap_het"] += 1
        elif "tạm đình chỉ" in kq.lower() or "tạm đình chỉ" in urg:
            st["tam_dinh_chi"] += 1
        elif any(k in kq.lower() for k in ("khởi tố", "không khởi tố", "đã nhập", "chuyển nơi khác", "không thụ lý")):
            st["da_giai_quyet"] += 1
        else:
            st["binh_thuong"] += 1

    sorted_units = sorted(unit_stats.keys(), key=lambda x: unit_stats[x]["total_tb"], reverse=True)
    return unit_stats, active_outcomes, sorted_units, tot_all


def build_officer_assignment_pivot_summary(records, is_tin_bao=False):
    """Tính toán ma trận phân công nghiệp vụ: Lãnh đạo Phòng / KSV x Khối lượng & Tiến độ giải quyết."""
    officer_stats = {}
    tot_all = len(records)
    for r in records:
        if is_tin_bao:
            ksv = str(r.get("ksv_chinh") or "").strip() or "Chưa phân công KSV"
            status = str(r.get("ket_qua_giai_quyet") or "").strip() or "Đang giải quyết"
            urg = str(r.get("muc_do_khan_cap") or "").lower()
        else:
            ksv = str(r.get("can_bo_giai_quyet") or "").strip() or "Chưa phân công KSV"
            status = str(r.get("trang_thai_vu_an") or r.get("giai_doan") or "").strip() or "Đang thụ lý"
            urg = str(r.get("muc_do_khan_cap") or "").lower()

        if ksv not in officer_stats:
            officer_stats[ksv] = {
                "total": 0,
                "qua_han": 0,
                "khan_cap": 0,
                "tam_dinh_chi": 0,
                "hoan_thanh": 0,
                "status_map": {}
            }
        st = officer_stats[ksv]
        st["total"] += 1
        st["status_map"][status] = st["status_map"].get(status, 0) + 1
        if "⚠️" in urg or "hết thời hiệu" in urg or "quá hạn" in urg:
            st["qua_han"] += 1
        elif "🔴" in urg or "khẩn cấp" in urg:
            st["khan_cap"] += 1
        elif "tạm đình chỉ" in status.lower() or "⏸" in urg:
            st["tam_dinh_chi"] += 1
        elif any(k in status.lower() for k in ("xét xử", "đình chỉ", "hoàn thành", "khởi tố", "không khởi tố")):
            st["hoan_thanh"] += 1

    sorted_officers = sorted(officer_stats.keys(), key=lambda x: officer_stats[x]["total"], reverse=True)
    return officer_stats, sorted_officers, tot_all


def create_simple_stage_pivot_sheet(wb, records, title="Pivot_Giai_Doan_VKS", index=None, report_title=None):
    """
    Tạo Sheet Pivot Table tự động theo mẫu bảng tổng hợp của VKS:
    - Tiêu đề gộp dòng 2-3: DANH SÁCH VỤ ÁN HẾT THỜI HIỆU TRUY CỨU TNHS CẬP NHẬT LÚC ...
    - Hàng: ĐƠN VỊ (Đơn vị VKS thụ lý)
    - Cột: Giai đoạn tố tụng (Điều tra, Truy tố, Sơ thẩm, Phúc thẩm...), TỔNG SỐ VỤ, Tỷ lệ %
    - Dòng cuối: TỔNG CỘNG
    """
    matrix, active_stages, sorted_units = build_pivot_stage_summary(records)
    tot_all = len(records)

    ws = wb.create_sheet(title=title, index=index)
    ws.views.sheetView[0].showGridLines = True

    NAVY = "07172C"
    NAVY_LIGHT = "0F2942"
    WHITE = "FFFFFF"
    BORDER_CLR = "CBD5E1"
    ZEBRA_BG = "F8FAFC"

    thin_border = Border(
        left=Side(style="thin", color=BORDER_CLR),
        right=Side(style="thin", color=BORDER_CLR),
        top=Side(style="thin", color=BORDER_CLR),
        bottom=Side(style="thin", color=BORDER_CLR),
    )
    double_bottom_border = Border(
        left=Side(style="thin", color=BORDER_CLR),
        right=Side(style="thin", color=BORDER_CLR),
        top=Side(style="thin", color=BORDER_CLR),
        bottom=Side(style="double", color="0F172A"),
    )

    col_gt = len(active_stages) + 2
    col_pct = col_gt + 1
    last_col_letter = get_column_letter(col_pct)

    # Dòng 2-3: Tiêu đề bảng tổng hợp gộp từ cột A đến cột cuối (Tỷ lệ %)
    now = datetime.now()
    if not report_title:
        time_str = f"{now.hour}H{now.minute:02d} NGÀY {now.strftime('%d/%m/%Y')}"
        report_title = f"DANH SÁCH VỤ ÁN HẾT THỜI HIỆU TRUY CỨU TNHS CẬP NHẬT LÚC {time_str}"

    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 22
    ws.merge_cells(f"A2:{last_col_letter}3")
    c_title = ws.cell(row=2, column=1, value=report_title)
    c_title.font = Font(name="Calibri", size=13, bold=True, color="000000")
    c_title.alignment = Alignment(horizontal="center", vertical="center")

    # Dòng 4: ĐƠN VỊ | [Stages...] | TỔNG SỐ VỤ | Tỷ lệ %
    ws.row_dimensions[4].height = 26
    c_rl = ws.cell(row=4, column=1, value="ĐƠN VỊ")
    c_rl.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c_rl.fill = PatternFill(start_color=NAVY_LIGHT, end_color=NAVY_LIGHT, fill_type="solid")
    c_rl.alignment = Alignment(horizontal="left", vertical="center")
    c_rl.border = thin_border

    for idx_s, s in enumerate(active_stages, 2):
        c = ws.cell(row=4, column=idx_s, value=s)
        c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border

    c_gt = ws.cell(row=4, column=col_gt, value="TỔNG SỐ VỤ")
    c_gt.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c_gt.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    c_gt.alignment = Alignment(horizontal="center", vertical="center")
    c_gt.border = thin_border

    c_pct_h = ws.cell(row=4, column=col_pct, value="Tỷ lệ %")
    c_pct_h.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c_pct_h.fill = PatternFill(start_color="047857", end_color="047857", fill_type="solid")
    c_pct_h.alignment = Alignment(horizontal="center", vertical="center")
    c_pct_h.border = thin_border

    # Data rows
    cur_row = 5
    for u_idx, u in enumerate(sorted_units, 1):
        ws.row_dimensions[cur_row].height = 22
        st = matrix[u]
        pct = (st["total_cases"] / tot_all * 100) if tot_all > 0 else 0.0
        bg_c = ZEBRA_BG if u_idx % 2 == 0 else WHITE

        c_u = ws.cell(row=cur_row, column=1, value=u)
        c_u.font = Font(name="Calibri", size=11)
        c_u.alignment = Alignment(horizontal="left", vertical="center")
        c_u.fill = PatternFill(start_color=bg_c, end_color=bg_c, fill_type="solid")
        c_u.border = thin_border

        for idx_s, s in enumerate(active_stages, 2):
            v = st["stages"].get(s, 0)
            c = ws.cell(row=cur_row, column=idx_s, value=v if v > 0 else "")
            c.font = Font(name="Calibri", size=11, bold=(v > 0))
            c.alignment = Alignment(horizontal="right", vertical="center")
            c.fill = PatternFill(start_color=bg_c, end_color=bg_c, fill_type="solid")
            c.border = thin_border

        c_tot = ws.cell(row=cur_row, column=col_gt, value=st["total_cases"])
        c_tot.font = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
        c_tot.alignment = Alignment(horizontal="right", vertical="center")
        c_tot.fill = PatternFill(start_color=bg_c, end_color=bg_c, fill_type="solid")
        c_tot.border = thin_border

        c_p = ws.cell(row=cur_row, column=col_pct, value=f"{pct:.1f}%")
        c_p.font = Font(name="Calibri", size=11)
        c_p.alignment = Alignment(horizontal="right", vertical="center")
        c_p.fill = PatternFill(start_color=bg_c, end_color=bg_c, fill_type="solid")
        c_p.border = thin_border

        cur_row += 1

    # TỔNG CỘNG row
    ws.row_dimensions[cur_row].height = 26
    c_gt_label = ws.cell(row=cur_row, column=1, value="TỔNG CỘNG")
    c_gt_label.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    c_gt_label.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    c_gt_label.alignment = Alignment(horizontal="left", vertical="center")
    c_gt_label.border = double_bottom_border

    for idx_s, s in enumerate(active_stages, 2):
        tot_s = sum(matrix[u]["stages"].get(s, 0) for u in sorted_units)
        c = ws.cell(row=cur_row, column=idx_s, value=tot_s)
        c.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
        c.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
        c.alignment = Alignment(horizontal="right", vertical="center")
        c.border = double_bottom_border

    c_all_gt = ws.cell(row=cur_row, column=col_gt, value=tot_all)
    c_all_gt.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    c_all_gt.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    c_all_gt.alignment = Alignment(horizontal="right", vertical="center")
    c_all_gt.border = double_bottom_border

    c_all_pct = ws.cell(row=cur_row, column=col_pct, value="100.0%")
    c_all_pct.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    c_all_pct.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    c_all_pct.alignment = Alignment(horizontal="right", vertical="center")
    c_all_pct.border = double_bottom_border

    # Adjust column widths
    ws.column_dimensions["A"].width = 38
    for idx_s in range(2, col_pct + 1):
        col_letter = get_column_letter(idx_s)
        ws.column_dimensions[col_letter].width = 16

    ws.freeze_panes = "B5"
    return ws


def create_case_risk_statute_pivot_sheet(wb, records, title="Ma_Tran_Rui_Ro_Thoi_Hieu", index=None):
    """
    Tạo Sheet Ma trận Rủi ro Thời hiệu truy cứu TNHS & Khẩn cấp dành cho Vụ án (Điều 230 BLTTHS & Điều 27 BLHS):
    Đơn vị VKS thụ lý x Phân loại mức độ khẩn cấp & thời hiệu.
    """
    unit_stats, _, sorted_units, tot_all = build_executive_master_pivot_summary(records)

    ws = wb.create_sheet(title=title, index=index)
    ws.views.sheetView[0].showGridLines = True

    # 1. Title Banner
    ws.merge_cells("A1:J1")
    t1 = ws["A1"]
    t1.value = "MA TRẬN GIÁM SÁT RỦI RO THỜI HIỆU TRUY CỨU TNHS & TIẾN ĐỘ VỤ ÁN (THEO Đ.230 BLTTHS & Đ.27 BLHS)"
    t1.font = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
    t1.fill = PatternFill(start_color="7F1D1D", end_color="7F1D1D", fill_type="solid")
    t1.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:J2")
    t2 = ws["A2"]
    t2.value = f"★ Báo cáo rà soát chuyên sâu các vụ án tạm đình chỉ, sắp hết hạn & hết thời hiệu - Toàn viện: {tot_all:,} vụ án ★"
    t2.font = Font(name="Segoe UI", size=10, bold=True, color="047857")
    t2.fill = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
    t2.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 24

    cell_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )
    double_bottom_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="double", color="0F172A")
    )

    risk_headers = [
        ("STT", 6), ("Đơn vị VKS thụ lý", 34), ("Tổng số vụ", 14),
        ("Hết thời hiệu ⚠️ (Đ.230)", 24), ("🔴 Khẩn cấp (≤7 ngày)", 20), ("🟠 Rất gấp (8-15 ngày)", 20),
        ("🟡 Sắp hết (16-30 ngày)", 20), ("🟢 Trong hạn (>30 ngày)", 20), ("⏸️ Tạm đình chỉ", 18), ("Tỷ lệ %", 12)
    ]
    ws.row_dimensions[4].height = 26
    for i, (h, w) in enumerate(risk_headers, start=1):
        c = ws.cell(row=4, column=i, value=h)
        c.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color="991B1B", end_color="991B1B", fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = cell_border
        ws.column_dimensions[get_column_letter(i)].width = w

    for u_idx, u in enumerate(sorted_units, start=1):
        curr_row = 4 + u_idx
        ws.row_dimensions[curr_row].height = 22
        st = unit_stats[u]
        pct = (st["total_cases"] / tot_all * 100) if tot_all else 0
        bg_row = "F8FAFC" if u_idx % 2 == 0 else "FFFFFF"
        vals = [
            u_idx, u, st["total_cases"],
            st["qua_han"] or "", st["khan_cap"] or "", st["rat_gap"] or "",
            st["sap_het"] or "", st["binh_thuong"] or "", st["tam_dinh_chi"] or "",
            f"{pct:.1f}%"
        ]
        for c_idx, val in enumerate(vals, start=1):
            cell_item = ws.cell(row=curr_row, column=c_idx, value=val)
            cell_item.border = cell_border
            align = "left" if c_idx == 2 else "center"
            cell_item.alignment = Alignment(horizontal=align, vertical="center")
            cell_item.fill = PatternFill(start_color=bg_row, end_color=bg_row, fill_type="solid")
            if c_idx == 4 and st["qua_han"] > 0:
                cell_item.fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
                cell_item.font = Font(name="Segoe UI", size=9, bold=True, color="991B1B")
            elif c_idx == 5 and st["khan_cap"] > 0:
                cell_item.fill = PatternFill(start_color="FEF2F2", end_color="FEF2F2", fill_type="solid")
                cell_item.font = Font(name="Segoe UI", size=9, bold=True, color="DC2626")

    tot_r_idx = 5 + len(sorted_units)
    ws.row_dimensions[tot_r_idx].height = 26
    tot_vals = [
        "∑", "TỔNG CỘNG TOÀN VIỆN", tot_all,
        sum(unit_stats[u]["qua_han"] for u in sorted_units),
        sum(unit_stats[u]["khan_cap"] for u in sorted_units),
        sum(unit_stats[u]["rat_gap"] for u in sorted_units),
        sum(unit_stats[u]["sap_het"] for u in sorted_units),
        sum(unit_stats[u]["binh_thuong"] for u in sorted_units),
        sum(unit_stats[u]["tam_dinh_chi"] for u in sorted_units),
        "100.0%"
    ]
    for c_idx, val in enumerate(tot_vals, start=1):
        cell_item = ws.cell(row=tot_r_idx, column=c_idx, value=val)
        cell_item.border = double_bottom_border
        cell_item.font = Font(name="Segoe UI", size=9, bold=True, color="0F172A")
        cell_item.fill = PatternFill(start_color="FCA5A5", end_color="FCA5A5", fill_type="solid")
        cell_item.alignment = Alignment(horizontal="left" if c_idx == 2 else "center", vertical="center")

    ws.freeze_panes = "C5"
    return ws


def create_executive_pivot_sheet(wb, records, title_prefix="BÁO CÁO TỔNG QUAN MA TRẬN PIVOT", title="Tong_Quan_Pivot_Lanh_Dao", index=None):
    """
    Tạo Sheet Báo cáo Tổng quan Pivot Table đa chiều phục vụ Lãnh đạo Viện / Lãnh đạo Phòng:
    - Bảng 1: Bảng Pivot gọn Đơn vị VKS × Giai đoạn tố tụng (Đúng y như mẫu người dùng yêu cầu).
    - Bảng 2: Ma trận đa chiều toàn diện 360° (Giai đoạn + Thời hiệu khẩn cấp + Phân loại tội danh + Số bị can).
    """
    unit_stats, active_stages, sorted_units, tot_all = build_executive_master_pivot_summary(records)

    ws = wb.create_sheet(title=title, index=index)
    ws.views.sheetView[0].showGridLines = True

    NAVY = "07172C"
    NAVY_LIGHT = "0F2942"
    BLUE = "0B2545"
    GOLD = "FBBF24"
    WHITE = "FFFFFF"
    BORDER_CLR = "CBD5E1"
    ZEBRA_BG = "F8FAFC"

    thin_border = Border(
        left=Side(style="thin", color=BORDER_CLR),
        right=Side(style="thin", color=BORDER_CLR),
        top=Side(style="thin", color=BORDER_CLR),
        bottom=Side(style="thin", color=BORDER_CLR),
    )
    double_bottom_border = Border(
        left=Side(style="thin", color=BORDER_CLR),
        right=Side(style="thin", color=BORDER_CLR),
        top=Side(style="thin", color=BORDER_CLR),
        bottom=Side(style="double", color="0F172A"),
    )

    # 1. Header Banner Lãnh Đạo
    ws.merge_cells("A1:R1")
    ws["A1"] = "VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH"
    ws["A1"].font = Font(name="Segoe UI", size=12, bold=True, color=GOLD)
    ws["A1"].fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:R2")
    ws["A2"] = f"{title_prefix.upper()}"
    ws["A2"].font = Font(name="Segoe UI", size=14, bold=True, color=WHITE)
    ws["A2"].fill = PatternFill(start_color=BLUE, end_color=BLUE, fill_type="solid")
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 34

    ws.merge_cells("A3:R3")
    ws["A3"] = f"★ Báo cáo số liệu quản lý điều hành & rà soát đối chiếu toàn Viện - Tổng số: {tot_all:,} vụ án ★"
    ws["A3"].font = Font(name="Segoe UI", size=10, bold=True, color="047857")
    ws["A3"].fill = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
    ws["A3"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 24

    # 2. BẢNG 1: PIVOT THEO ĐƠN VỊ VKS VÀ GIAI ĐOẠN TỐ TỤNG (ĐÚNG YÊU CẦU NGƯỜI DÙNG)
    b1_last_col = get_column_letter(len(active_stages) + 4)
    ws.merge_cells(f"A5:{b1_last_col}5")
    ws["A5"] = "🏛️ BẢNG 1: PIVOT THEO ĐƠN VỊ VKS VÀ TIẾN TRÌNH GIAI ĐOẠN TỐ TỤNG (ĐIỀU TRA - TRUY TỐ - SƠ THẨM)"
    ws["A5"].font = Font(name="Segoe UI", size=10, bold=True, color="1E3A8A")
    ws["A5"].fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
    ws["A5"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[5].height = 24

    b1_headers = ["STT", "Đơn vị VKS thụ lý (Row Labels)"] + active_stages + ["Grand Total", "Tỷ lệ %"]
    ws.row_dimensions[6].height = 26
    for c_idx, h in enumerate(b1_headers, 1):
        cell = ws.cell(row=6, column=c_idx, value=h)
        cell.font = Font(name="Segoe UI", size=9, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=NAVY_LIGHT, end_color=NAVY_LIGHT, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    cur_row = 7
    for u_idx, u_name in enumerate(sorted_units, 1):
        ws.row_dimensions[cur_row].height = 22
        st = unit_stats[u_name]
        pct = (st["total_cases"] / tot_all * 100) if tot_all > 0 else 0.0
        bg_row = ZEBRA_BG if u_idx % 2 == 0 else WHITE

        ws.cell(row=cur_row, column=1, value=u_idx).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=cur_row, column=2, value=u_name).alignment = Alignment(horizontal="left", vertical="center")

        col_pos = 3
        for s in active_stages:
            val = st["stages"].get(s, 0)
            c = ws.cell(row=cur_row, column=col_pos, value=val if val > 0 else "")
            c.alignment = Alignment(horizontal="center", vertical="center")
            if val > 0:
                c.font = Font(name="Segoe UI", size=9, bold=True, color="0F172A")
            col_pos += 1

        c_tot = ws.cell(row=cur_row, column=col_pos, value=st["total_cases"])
        c_tot.alignment = Alignment(horizontal="center", vertical="center")
        c_tot.font = Font(name="Segoe UI", size=9, bold=True, color="1E3A8A")

        c_pct = ws.cell(row=cur_row, column=col_pos + 1, value=f"{pct:.1f}%")
        c_pct.alignment = Alignment(horizontal="center", vertical="center")

        for c_i in range(1, len(b1_headers) + 1):
            cell = ws.cell(row=cur_row, column=c_i)
            cell.border = thin_border
            if cell.fill.fill_type is None:
                cell.fill = PatternFill(start_color=bg_row, end_color=bg_row, fill_type="solid")

        cur_row += 1

    # Dòng Grand Total Bảng 1
    ws.row_dimensions[cur_row].height = 26
    ws.cell(row=cur_row, column=1, value="∑").alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=cur_row, column=2, value="Grand Total (TỔNG CỘNG CHUNG)").alignment = Alignment(horizontal="center", vertical="center")

    col_pos = 3
    for s in active_stages:
        s_tot = sum(unit_stats[u]["stages"].get(s, 0) for u in sorted_units)
        ws.cell(row=cur_row, column=col_pos, value=s_tot).alignment = Alignment(horizontal="center", vertical="center")
        col_pos += 1

    ws.cell(row=cur_row, column=col_pos, value=tot_all).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=cur_row, column=col_pos + 1, value="100.0%").alignment = Alignment(horizontal="center", vertical="center")

    for c_i in range(1, len(b1_headers) + 1):
        cell = ws.cell(row=cur_row, column=c_i)
        cell.font = Font(name="Segoe UI", size=10, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
        cell.border = double_bottom_border

    # 3. BẢNG 2: MA TRẬN TỔNG HỢP TOÀN DIỆN SỐ LIỆU ĐIỀU HÀNH DÀNH CHO LÃNH ĐẠO (ĐIỀU HÀNH 360°)
    cur_row += 3
    b2_title_row = cur_row
    ws.merge_cells(f"A{b2_title_row}:R{b2_title_row}")
    ws[f"A{b2_title_row}"] = "⚖️ BẢNG 2: MA TRẬN ĐA CHIỀU TOÀN DIỆN PHỤC VỤ CÔNG TÁC QUẢN LÝ, CHỈ ĐẠO & ĐIỀU HÀNH (DÀNH CHO LÃNH ĐẠO)"
    ws[f"A{b2_title_row}"].font = Font(name="Segoe UI", size=11, bold=True, color="991B1B")
    ws[f"A{b2_title_row}"].fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    ws[f"A{b2_title_row}"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[b2_title_row].height = 26

    # Header Cấp 1 (Group Banner Bảng 2)
    cur_row += 1
    g1_row = cur_row
    ws.row_dimensions[g1_row].height = 22

    GROUPS_B2 = [
        ("A", "B", "1. ĐƠN VỊ VKS THỤ LÝ", "0F2942", "FFFFFF"),
        ("C", "E", "2. TIẾN TRÌNH GIAI ĐOẠN TỐ TỤNG", "1E3A8A", "DBEAFE"),
        ("F", "K", "3. CẢNH BÁO THỜI HIỆU & KHẨN CẤP (ĐẶC BIỆT CHÚ Ý)", "991B1B", "FEE2E2"),
        ("L", "O", "4. PHÂN LOẠI TÍNH CHẤT TỘI PHẠM", "166534", "DCFCE7"),
        ("P", "R", "5. TỔNG HỢP BỊ CAN & CƠ CẤU", "581C87", "F3E8FF"),
    ]
    for c_start, c_end, g_title, bg_c, fg_c in GROUPS_B2:
        range_str = f"{c_start}{g1_row}:{c_end}{g1_row}"
        ws.merge_cells(range_str)
        top_cell = ws[f"{c_start}{g1_row}"]
        top_cell.value = g_title
        top_cell.font = Font(name="Segoe UI", size=9, bold=True, color=fg_c)
        top_cell.fill = PatternFill(start_color=bg_c, end_color=bg_c, fill_type="solid")
        top_cell.alignment = Alignment(horizontal="center", vertical="center")

    # Header Cấp 2 (Columns Header Bảng 2)
    cur_row += 1
    h2_row = cur_row
    ws.row_dimensions[h2_row].height = 26

    COLS_B2 = [
        ("STT", 6, "center", "0F2942"),
        ("Đơn vị VKS thụ lý", 36, "left", "0F2942"),
        ("Điều tra", 11, "center", "1E3A8A"),
        ("Truy tố", 11, "center", "1E3A8A"),
        ("Sơ thẩm", 11, "center", "1E3A8A"),
        ("Quá hạn ⚠️", 13, "center", "991B1B"),
        ("Khẩn cấp 🔴", 13, "center", "991B1B"),
        ("Rất gấp 🟠", 12, "center", "991B1B"),
        ("Sắp hết 🟡", 12, "center", "991B1B"),
        ("Trong hạn 🟢", 13, "center", "991B1B"),
        ("Tạm đình chỉ ⏸️", 14, "center", "991B1B"),
        ("Ít NT", 11, "center", "166534"),
        ("Nghiêm trọng", 13, "center", "166534"),
        ("Rất NT", 11, "center", "166534"),
        ("Đặc biệt NT", 13, "center", "166534"),
        ("Tổng số vụ", 12, "center", "581C87"),
        ("Tổng bị can", 12, "center", "581C87"),
        ("Tỷ lệ %", 11, "center", "581C87"),
    ]

    for c_idx, (col_title, col_w, col_align, col_bg) in enumerate(COLS_B2, 1):
        cell = ws.cell(row=h2_row, column=c_idx, value=col_title)
        cell.font = Font(name="Segoe UI", size=9, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=col_bg, end_color=col_bg, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        col_letter = get_column_letter(c_idx)
        if ws.column_dimensions[col_letter].width is None or ws.column_dimensions[col_letter].width < col_w:
            ws.column_dimensions[col_letter].width = col_w

    # Điền dữ liệu Bảng 2
    cur_row += 1
    fill_red_light = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    font_red_bold = Font(name="Segoe UI", size=9, bold=True, color="991B1B")
    fill_amber_light = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    font_amber_bold = Font(name="Segoe UI", size=9, bold=True, color="92400E")

    for u_idx, u_name in enumerate(sorted_units, 1):
        ws.row_dimensions[cur_row].height = 22
        st = unit_stats[u_name]
        pct = (st["total_cases"] / tot_all * 100) if tot_all > 0 else 0.0
        bg_row = ZEBRA_BG if u_idx % 2 == 0 else WHITE

        row_vals = [
            u_idx,
            u_name,
            st["stages"].get("Điều tra", 0) or "",
            st["stages"].get("Truy tố", 0) or "",
            st["stages"].get("Sơ thẩm", 0) or "",
            st["qua_han"] or "",
            st["khan_cap"] or "",
            st["rat_gap"] or "",
            st["sap_het"] or "",
            st["binh_thuong"] or "",
            st["tam_dinh_chi"] or "",
            st["toi_it"] or "",
            st["toi_nt"] or "",
            st["toi_rnt"] or "",
            st["toi_dbnt"] or "",
            st["total_cases"],
            st["total_defendants"],
            f"{pct:.1f}%",
        ]

        for c_idx, val in enumerate(row_vals, 1):
            cell = ws.cell(row=cur_row, column=c_idx, value=val)
            align = "left" if c_idx == 2 else "center"
            cell.alignment = Alignment(horizontal=align, vertical="center")
            cell.border = thin_border
            cell.fill = PatternFill(start_color=bg_row, end_color=bg_row, fill_type="solid")
            cell.font = Font(name="Segoe UI", size=9)

            if c_idx == 6 and val: # Quá hạn
                cell.fill = fill_red_light
                cell.font = font_red_bold
            elif c_idx == 7 and val: # Khẩn cấp
                cell.fill = fill_amber_light
                cell.font = font_amber_bold
            elif c_idx in (16, 17): # Tổng vụ, tổng bị can
                cell.font = Font(name="Segoe UI", size=9, bold=True, color="0F172A")

        cur_row += 1

    # Dòng Grand Total Bảng 2
    ws.row_dimensions[cur_row].height = 26
    ws.cell(row=cur_row, column=1, value="∑").alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=cur_row, column=2, value="Grand Total (TỔNG CỘNG TOÀN VIỆN)").alignment = Alignment(horizontal="center", vertical="center")

    tot_vals = [
        sum(unit_stats[u]["stages"].get("Điều tra", 0) for u in sorted_units),
        sum(unit_stats[u]["stages"].get("Truy tố", 0) for u in sorted_units),
        sum(unit_stats[u]["stages"].get("Sơ thẩm", 0) for u in sorted_units),
        sum(unit_stats[u]["qua_han"] for u in sorted_units),
        sum(unit_stats[u]["khan_cap"] for u in sorted_units),
        sum(unit_stats[u]["rat_gap"] for u in sorted_units),
        sum(unit_stats[u]["sap_het"] for u in sorted_units),
        sum(unit_stats[u]["binh_thuong"] for u in sorted_units),
        sum(unit_stats[u]["tam_dinh_chi"] for u in sorted_units),
        sum(unit_stats[u]["toi_it"] for u in sorted_units),
        sum(unit_stats[u]["toi_nt"] for u in sorted_units),
        sum(unit_stats[u]["toi_rnt"] for u in sorted_units),
        sum(unit_stats[u]["toi_dbnt"] for u in sorted_units),
        tot_all,
        sum(unit_stats[u]["total_defendants"] for u in sorted_units),
        "100.0%",
    ]

    for c_i, val in enumerate(tot_vals, 3):
        ws.cell(row=cur_row, column=c_i, value=val).alignment = Alignment(horizontal="center", vertical="center")

    for c_i in range(1, 19):
        cell = ws.cell(row=cur_row, column=c_i)
        cell.font = Font(name="Segoe UI", size=10, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
        cell.border = double_bottom_border

    ws.freeze_panes = "C7"
    return ws


def export_case_records_to_excel(records, out_filepath):
    """Xuất file Excel Sổ vụ án chuyên nghiệp với định dạng ngày tháng chuẩn Excel,
    phân nhóm tiêu đề 2 cấp, tô màu khẩn cấp, bật AutoFilter và cố định dòng cột.
    Bao gồm 4 Sheet:
    - Sheet 1: Tong_Quan_Pivot_Lanh_Dao (Báo cáo Đa Chiều Toàn Diện Lãnh Đạo)
    - Sheet 2: Pivot_Giai_Doan_VKS (Ma Trận Pivot Đơn vị VKS x Giai đoạn tố tụng)
    - Sheet 3: Danh_Sach_Vu_An_Chuan_Hoa (Sổ 26 cột nghiệp vụ chuẩn hóa)
    - Sheet 4: Tong_Hop_Pivot_Table (Bảng phân tích khẩn cấp thời hiệu & tội phạm)
    """
    for rec in records:
        normalize_case_stage_and_status(rec)

    wb = openpyxl.Workbook()
    # Tạo Sheet 1, Sheet 2, Sheet 3: Báo cáo Tổng quan Pivot Table & Rủi ro Thời hiệu phục vụ Lãnh đạo
    create_executive_pivot_sheet(wb, records, title_prefix="BÁO CÁO TỔNG QUAN MA TRẬN PIVOT SỔ VỤ ÁN QLAHS", title="Tong_Quan_Pivot_Lanh_Dao", index=0)
    create_simple_stage_pivot_sheet(wb, records, title="Pivot_Giai_Doan_VKS", index=1)
    create_case_risk_statute_pivot_sheet(wb, records, title="Ma_Tran_Rui_Ro_Thoi_Hieu", index=2)

    ws1 = wb.create_sheet("Danh_Sach_Vu_An_Chuan_Hoa", index=3)
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])
    ws1.views.sheetView[0].showGridLines = True

    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )
    double_bottom_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="double", color="0F172A"),
    )

    # 26 Cột chuẩn hóa theo đúng cấu trúc nghiệp vụ
    COLUMNS_26 = [
        ("stt", "STT", 6, "center", "0F2942"),
        ("ma_vu_an", "Mã vụ án", 28, "center", "0F2942"),
        ("ten_vu_an", "Tên vụ án / Nội dung / QĐ TĐC", 38, "left", "0F2942"),
        ("tinh_thanh", "Tỉnh / Thành phố", 18, "center", "0F2942"),
        ("so_qd_khoi_to", "Số QĐ khởi tố", 14, "center", "1E3A8A"),
        ("ngay_qd_khoi_to", "Ngày ra QĐ khởi tố", 16, "center", "1E3A8A"),
        ("don_vi_khoi_to", "Đơn vị ra QĐ khởi tố", 30, "left", "1E3A8A"),
        ("so_thu_ly", "Số thụ lý", 14, "center", "0E7490"),
        ("ngay_thu_ly", "Ngày thụ lý", 16, "center", "0E7490"),
        ("loai_thu_ly", "Loại thụ lý", 16, "center", "0E7490"),
        ("vks_thu_ly", "Đơn vị VKS thụ lý", 34, "left", "0E7490"),
        ("ngay_tiep_nhan", "Ngày tiếp nhận hồ sơ", 16, "center", "991B1B"),
        ("thoi_han_giai_quyet", "Thời hạn giải quyết", 16, "center", "991B1B"),
        ("ngay_het_thoi_hieu", "Ngày hết thời hiệu TNHS", 18, "center", "991B1B"),
        ("so_ngay_con_lai", "Số ngày còn lại", 15, "center", "991B1B"),
        ("muc_do_khan_cap", "Mức độ khẩn cấp", 24, "center", "991B1B"),
        ("canh_bao_thoi_hieu", "Cảnh báo thời hiệu", 26, "center", "991B1B"),
        ("bi_can_dau_vu", "Bị can đầu vụ & CCCD", 26, "left", "4C1D95"),
        ("tong_bi_can", "Tổng số bị can", 14, "center", "4C1D95"),
        ("dieu_luat", "Thông tin tội danh / Điều luật", 36, "left", "4C1D95"),
        ("loai_toi", "Phân loại tội", 18, "center", "4C1D95"),
        ("can_bo_thu_ly", "Cán bộ thụ lý", 18, "left", "334155"),
        ("can_bo_giai_quyet", "Cán bộ giải quyết (KSV)", 28, "left", "334155"),
        ("giai_doan", "Giai đoạn tố tụng", 16, "center", "334155"),
        ("trang_thai_vu_an", "Trạng thái vụ án", 22, "center", "334155"),
        ("trang_thai_bieu_mau", "Trạng thái biểu mẫu", 22, "left", "334155"),
    ]

    # Banner Header (Rows 1-3)
    ws1.merge_cells("A1:Z1")
    ws1["A1"] = "VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH"
    ws1["A1"].font = Font(name="Segoe UI", size=12, bold=True, color="FBBF24")
    ws1["A1"].fill = PatternFill(start_color="07172C", end_color="07172C", fill_type="solid")
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 28

    ws1.merge_cells("A2:Z2")
    ws1["A2"] = "BẢNG CHUẨN HÓA DỮ LIỆU SỔ THỤ LÝ & THEO DÕI THỜI HIỆU TRUY CỨU TNHS"
    ws1["A2"].font = Font(name="Segoe UI", size=15, bold=True, color="FFFFFF")
    ws1["A2"].fill = PatternFill(start_color="0B2545", end_color="0B2545", fill_type="solid")
    ws1["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[2].height = 34

    ws1.merge_cells("A3:Z3")
    ws1["A3"] = f"★ Tự động chuẩn hóa dữ liệu QLAHS - Tổng số: {len(records)} vụ án - 100% Bảo mật Offline ★"
    ws1["A3"].font = Font(name="Segoe UI", size=10, bold=True, color="991B1B")
    ws1["A3"].fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    ws1["A3"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[3].height = 24

    # Row 4: Group Header Banners
    GROUP_BANNERS = [
        ("A4:D4", "1. THÔNG TIN VỤ ÁN", "0F2942", "F8FAFC"),
        ("E4:G4", "2. QUYẾT ĐỊNH KHỞI TỐ", "1E3A8A", "DBEAFE"),
        ("H4:K4", "3. THỤ LÝ KIỂM SÁT", "0E7490", "CFFAFE"),
        ("L4:Q4", "4. TIẾN ĐỘ & CẢNH BÁO THỜI HIỆU TNHS (ĐẶC BIỆT CHÚ Ý)", "991B1B", "FEE2E2"),
        ("R4:U4", "5. BỊ CAN & TỘI DANH", "4C1D95", "EDE9FE"),
        ("V4:Z4", "6. PHÂN CÔNG & TRẠNG THÁI TỐ TỤNG", "334155", "F1F5F9"),
    ]
    ws1.row_dimensions[4].height = 24
    for range_str, g_title, bg_color, text_color in GROUP_BANNERS:
        ws1.merge_cells(range_str)
        top_cell = ws1[range_str.split(":")[0]]
        top_cell.value = g_title
        top_cell.font = Font(name="Segoe UI", size=10, bold=True, color=text_color)
        top_cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
        top_cell.alignment = Alignment(horizontal="center", vertical="center")
        for row_c in ws1[range_str]:
            for c in row_c:
                c.border = thin_border
                c.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")

    # Row 5: Column Headers
    ws1.row_dimensions[5].height = 32
    for col_idx, (k, title, width, align, bg_color) in enumerate(COLUMNS_26, 1):
        cell = ws1.cell(row=5, column=col_idx, value=title)
        cell.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        col_letter = get_column_letter(col_idx)
        ws1.column_dimensions[col_letter].width = width

    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_slate = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    fill_urgent_red = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fill_urgent_amber = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    fill_urgent_yellow = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")

    font_urgent_red = Font(name="Segoe UI", size=9, bold=True, color="991B1B")
    font_urgent_amber = Font(name="Segoe UI", size=9, bold=True, color="92400E")
    font_urgent_yellow = Font(name="Segoe UI", size=9, bold=True, color="854D0E")
    font_regular = Font(name="Segoe UI", size=9, color="0F172A")
    font_bold = Font(name="Segoe UI", size=9, bold=True, color="0F172A")
    font_date = Font(name="Segoe UI", size=9, bold=True, color="1E293B")

    date_field_keys = {"ngay_qd_khoi_to", "ngay_thu_ly", "ngay_tiep_nhan", "thoi_han_giai_quyet", "ngay_het_thoi_hieu"}
    today_date = date.today()

    for r_idx, rec in enumerate(records, 6):
        normalize_case_stage_and_status(rec)
        ws1.row_dimensions[r_idx].height = 24
        base_fill = fill_white if r_idx % 2 == 0 else fill_slate

        d_het_han = parse_date_safely(rec.get("ngay_het_thoi_hieu"))
        days_left = (d_het_han - today_date).days if d_het_han else None

        if days_left is not None:
            if days_left <= 7:
                urgency_fill = fill_urgent_red
                urgency_font = font_urgent_red
            elif days_left <= 15:
                urgency_fill = fill_urgent_amber
                urgency_font = font_urgent_amber
            else:
                urgency_fill = fill_urgent_yellow
                urgency_font = font_urgent_yellow
        else:
            urgency_fill = base_fill
            urgency_font = font_regular

        for c_idx, (k, _, _, align, _) in enumerate(COLUMNS_26, 1):
            cell = ws1.cell(row=r_idx, column=c_idx)
            val = rec.get(k, "")
            cell.border = thin_border

            if k in date_field_keys:
                d_val = parse_date_safely(val)
                if d_val:
                    cell.value = d_val
                    cell.number_format = 'DD/MM/YYYY'
                else:
                    cell.value = str(val or "")
                    cell.number_format = '@'
                cell.font = font_date
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if k == "ngay_het_thoi_hieu":
                    cell.fill = urgency_fill
                    cell.font = urgency_font
                else:
                    cell.fill = base_fill
            elif k in ("so_ngay_con_lai", "muc_do_khan_cap", "canh_bao_thoi_hieu"):
                val_str = str(val or "")
                cell.value = val_str
                cell.number_format = '@'
                cell.alignment = Alignment(horizontal="center", vertical="center")
                val_low = val_str.lower()
                if "⏸" in val_str or "tạm đình chỉ" in val_low:
                    cell.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
                    cell.font = Font(name="Segoe UI", size=9, bold=True, color="475569")
                elif "⏹" in val_str or "đình chỉ" in val_low:
                    cell.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
                    cell.font = Font(name="Segoe UI", size=9, bold=True, color="64748B")
                elif "✔" in val_str or "hoàn thành" in val_low or "đã giải quyết" in val_low:
                    cell.fill = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
                    cell.font = Font(name="Segoe UI", size=9, bold=True, color="047857")
                elif "⚖" in val_str or "đã xét xử" in val_low:
                    cell.fill = PatternFill(start_color="F5F3FF", end_color="F5F3FF", fill_type="solid")
                    cell.font = Font(name="Segoe UI", size=9, bold=True, color="581C87")
                else:
                    cell.fill = urgency_fill
                    cell.font = urgency_font
            elif k in ("loai_toi",):
                cell.value = str(val or "")
                cell.number_format = '@'
                cell.alignment = Alignment(horizontal="center", vertical="center")
                lt_low = str(val).lower()
                if "rất nghiêm trọng" in lt_low or "đặc biệt" in lt_low:
                    cell.fill = fill_urgent_red
                    cell.font = font_urgent_red
                elif "nghiêm trọng" in lt_low:
                    cell.fill = fill_urgent_amber
                    cell.font = font_urgent_amber
                elif "ít nghiêm trọng" in lt_low:
                    cell.fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
                    cell.font = Font(name="Segoe UI", size=9, bold=True, color="1D4ED8")
                else:
                    cell.fill = base_fill
                    cell.font = font_regular
            else:
                cell.value = val
                cell.fill = base_fill
                cell.font = font_bold if k in ("ma_vu_an", "stt", "tong_bi_can") else font_regular
                cell.alignment = Alignment(horizontal=align, vertical="center")
                if k in ("stt", "tong_bi_can") and isinstance(val, int):
                    cell.number_format = '#,##0'
                else:
                    cell.number_format = '@'

    # AutoFilter on row 5 and FreezePanes at C6
    ws1.auto_filter.ref = f"A5:Z{ws1.max_row}"
    ws1.freeze_panes = "C6"

    # Sheet 5: Bảng Pivot Table Tổng Hợp & Đối Soát Thời Hiệu
    ws2 = wb.create_sheet("Tong_Hop_Pivot_Table", index=4)
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:G1")
    ws2["A1"] = "BẢNG TỔNG HỢP & MA TRẬN ĐỐI SOÁT THỜI HIỆU TRUY CỨU TNHS THEO ĐƠN VỊ VKS"
    ws2["A1"].font = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
    ws2["A1"].fill = PatternFill(start_color="0B2545", end_color="0B2545", fill_type="solid")
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 32

    # Phần 1: Bảng theo dõi khẩn cấp thời hiệu
    ws2.merge_cells("A3:F3")
    ws2["A3"] = "PHẦN 1: THEO DÕI MỨC ĐỘ KHẨN CẤP THỜI HIỆU TNHS (TÍNH ĐẾN HÔM NAY)"
    ws2["A3"].font = Font(name="Segoe UI", size=10, bold=True, color="991B1B")
    ws2["A3"].fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    ws2["A3"].alignment = Alignment(horizontal="left", vertical="center")
    ws2.row_dimensions[3].height = 24

    headers_u = ["STT", "Đơn vị VKS thụ lý", "Tổng số vụ", "🔴 Khẩn cấp (≤ 7 ngày)", "🟠 Rất gấp (8-15 ngày)", "🟡 Sắp hết hạn (16-30 ngày)", "Tỷ lệ %"]
    ws2.row_dimensions[4].height = 28
    for c_i, h in enumerate(headers_u, 1):
        cell = ws2.cell(row=4, column=c_i, value=h)
        cell.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="991B1B" if c_i in (4, 5, 6) else "0F2942", end_color="991B1B" if c_i in (4, 5, 6) else "0F2942", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    unit_stats = {}
    for r in records:
        u = str(r.get("vks_thu_ly") or "").strip() or "Chưa rõ VKS"
        if u not in unit_stats:
            unit_stats[u] = {"total": 0, "under_7": 0, "8_to_15": 0, "16_to_30": 0, "crimes": {}}
        unit_stats[u]["total"] += 1

        d_h = parse_date_safely(r.get("ngay_het_thoi_hieu"))
        diff = (d_h - today_date).days if d_h else 999
        if diff <= 7:
            unit_stats[u]["under_7"] += 1
        elif diff <= 15:
            unit_stats[u]["8_to_15"] += 1
        else:
            unit_stats[u]["16_to_30"] += 1

        lt = str(r.get("loai_toi") or "").strip() or "Chưa phân loại"
        unit_stats[u]["crimes"][lt] = unit_stats[u]["crimes"].get(lt, 0) + 1

    sorted_units = sorted(unit_stats.keys(), key=lambda x: unit_stats[x]["under_7"], reverse=True)
    tot_all = len(records)
    cur_r = 5

    for idx, u in enumerate(sorted_units, 1):
        st = unit_stats[u]
        pct = st["total"] / tot_all * 100 if tot_all else 0
        ws2.row_dimensions[cur_r].height = 22
        r_f = fill_white if idx % 2 == 0 else fill_slate
        row_vals = [idx, u, st["total"], st["under_7"], st["8_to_15"], st["16_to_30"], f"{pct:.1f}%"]
        for c_i, val in enumerate(row_vals, 1):
            c = ws2.cell(row=cur_r, column=c_i, value=val)
            c.border = thin_border
            c.font = font_regular
            c.fill = r_f
            if c_i == 4 and val > 0:
                c.fill = fill_urgent_red
                c.font = font_urgent_red
            elif c_i == 5 and val > 0:
                c.fill = fill_urgent_amber
                c.font = font_urgent_amber
            elif c_i == 6 and val > 0:
                c.fill = fill_urgent_yellow
                c.font = font_urgent_yellow
            c.alignment = Alignment(horizontal="left" if c_i == 2 else "center", vertical="center")
        cur_r += 1

    # Totals Row Part 1
    tot_under_7 = sum(unit_stats[u]["under_7"] for u in sorted_units)
    tot_8_to_15 = sum(unit_stats[u]["8_to_15"] for u in sorted_units)
    tot_16_to_30 = sum(unit_stats[u]["16_to_30"] for u in sorted_units)
    ws2.row_dimensions[cur_r].height = 26
    tot_vals = ["∑", "TỔNG CỘNG CHUNG", tot_all, tot_under_7, tot_8_to_15, tot_16_to_30, "100.0%"]
    for c_i, val in enumerate(tot_vals, 1):
        c = ws2.cell(row=cur_r, column=c_i, value=val)
        c.fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
        c.font = Font(name="Segoe UI", size=9, bold=True, color="991B1B")
        c.border = double_bottom_border
        c.alignment = Alignment(horizontal="left" if c_i == 2 else "center", vertical="center")
    cur_r += 3

    # Phần 2: Bảng phân loại mức độ tội phạm
    ws2.merge_cells(f"A{cur_r}:F{cur_r}")
    ws2[f"A{cur_r}"] = "PHẦN 2: ĐỐI SOÁT PHÂN LOẠI MỨC ĐỘ NGHIÊM TRỌNG CỦA TỘI PHẠM THEO ĐƠN VỊ VKS"
    ws2[f"A{cur_r}"].font = Font(name="Segoe UI", size=10, bold=True, color="1E3A8A")
    ws2[f"A{cur_r}"].fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
    ws2[f"A{cur_r}"].alignment = Alignment(horizontal="left", vertical="center")
    ws2.row_dimensions[cur_r].height = 24
    cur_r += 1

    headers_c = ["STT", "Đơn vị VKS thụ lý", "Tổng số vụ", "Ít nghiêm trọng", "Nghiêm trọng", "Rất nghiêm trọng", "Tỷ lệ %"]
    ws2.row_dimensions[cur_r].height = 28
    for c_i, h in enumerate(headers_c, 1):
        cell = ws2.cell(row=cur_r, column=c_i, value=h)
        cell.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    cur_r += 1

    sorted_units_name = sorted(unit_stats.keys(), key=lambda x: unit_stats[x]["total"], reverse=True)
    for idx, u in enumerate(sorted_units_name, 1):
        st = unit_stats[u]
        pct = st["total"] / tot_all * 100 if tot_all else 0
        ws2.row_dimensions[cur_r].height = 22
        r_f = fill_white if idx % 2 == 0 else fill_slate
        it_nt = st["crimes"].get("Ít nghiêm trọng", 0)
        nt = st["crimes"].get("Nghiêm trọng", 0)
        rat_nt = st["crimes"].get("Rất nghiêm trọng", 0)
        row_vals = [idx, u, st["total"], it_nt, nt, rat_nt, f"{pct:.1f}%"]
        for c_i, val in enumerate(row_vals, 1):
            c = ws2.cell(row=cur_r, column=c_i, value=val)
            c.border = thin_border
            c.font = font_regular
            c.fill = r_f
            c.alignment = Alignment(horizontal="left" if c_i == 2 else "center", vertical="center")
        cur_r += 1

    # Totals Row Part 2
    tot_it_nt = sum(unit_stats[u]["crimes"].get("Ít nghiêm trọng", 0) for u in sorted_units_name)
    tot_nt = sum(unit_stats[u]["crimes"].get("Nghiêm trọng", 0) for u in sorted_units_name)
    tot_rat_nt = sum(unit_stats[u]["crimes"].get("Rất nghiêm trọng", 0) for u in sorted_units_name)
    ws2.row_dimensions[cur_r].height = 26
    tot_vals_2 = ["∑", "TỔNG CỘNG CHUNG", tot_all, tot_it_nt, tot_nt, tot_rat_nt, "100.0%"]
    for c_i, val in enumerate(tot_vals_2, 1):
        c = ws2.cell(row=cur_r, column=c_i, value=val)
        c.fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
        c.font = Font(name="Segoe UI", size=9, bold=True, color="B45309")
        c.border = double_bottom_border
        c.alignment = Alignment(horizontal="left" if c_i == 2 else "center", vertical="center")

    ws2.column_dimensions["A"].width = 8
    ws2.column_dimensions["B"].width = 38
    ws2.column_dimensions["C"].width = 16
    ws2.column_dimensions["D"].width = 24
    ws2.column_dimensions["E"].width = 24
    ws2.column_dimensions["F"].width = 26
    ws2.column_dimensions["G"].width = 14

    # Bổ sung Biểu đồ Excel trực quan (Data Charts)
    if BarChart and Reference and sorted_units:
        try:
            chart_case = BarChart()
            chart_case.type = "col"
            chart_case.style = 10
            chart_case.title = "BIỂU ĐỒ THEO DÕI SỐ LƯỢNG VỤ ÁN THEO ĐƠN VỊ VKS"
            chart_case.y_axis.title = "Số lượng vụ án"
            chart_case.x_axis.title = "Đơn vị VKS thụ lý"
            data_c = Reference(ws2, min_col=3, min_row=4, max_row=4 + len(sorted_units))
            cats_c = Reference(ws2, min_col=2, min_row=5, max_row=4 + len(sorted_units))
            chart_case.add_data(data_c, titles_from_data=True)
            chart_case.set_categories(cats_c)
            chart_case.width = 18
            chart_case.height = 10
            ws2.add_chart(chart_case, "I4")
        except Exception:
            pass

    wb.active = 0
    wb.save(out_filepath)




def export_reconciliation_excel(records, out_filepath):
    """Xuất file Excel đối chiếu & rà soát dữ liệu 29 cột chuyên nghiệp,
    bao gồm 5 Sheet:
    - Sheet 1: Tong_Quan_Pivot_Lanh_Dao (Báo cáo Đa Chiều Toàn Diện Lãnh Đạo)
    - Sheet 2: Pivot_Giai_Doan_VKS (Ma Trận Pivot Đơn vị VKS x Giai đoạn tố tụng)
    - Sheet 3: Danh_Sach_Hop_Nhat_Chuan_Hoa (Dữ liệu hợp nhất 29 cột nghiệp vụ)
    - Sheet 4: Bao_Cao_Ra_Soat_Doi_Chieu (Báo cáo phân công Lãnh đạo & đối chiếu)
    - Sheet 5: Tong_Hop_Pivot_Table (Bảng tổng hợp theo đơn vị VKS và loại tội)
    """
    # Chuẩn hóa sạch toàn bộ tên cán bộ & lãnh đạo (xóa mọi tiền tố PVT, Đ/c, KSV...) ngay từ đầu
    for rec in records:
        normalize_case_stage_and_status(rec)
        if not rec.get("lanh_dao_vien") and not rec.get("lanh_dao_phong"):
            ld_v, ld_p, ksv, cb_tl = split_officers(rec.get("can_bo_giai_quyet", ""), rec.get("can_bo_thu_ly", ""))
            rec["lanh_dao_vien"] = ld_v
            rec["lanh_dao_phong"] = ld_p
            rec["can_bo_giai_quyet"] = ksv
            rec["can_bo_thu_ly"] = cb_tl
        else:
            if rec.get("lanh_dao_vien"):
                rec["lanh_dao_vien"] = "; ".join(filter(None, [clean_person_name(x) for x in str(rec["lanh_dao_vien"]).split(";")]))
            if rec.get("lanh_dao_phong"):
                rec["lanh_dao_phong"] = "; ".join(filter(None, [clean_person_name(x) for x in str(rec["lanh_dao_phong"]).split(";")]))
            if rec.get("can_bo_giai_quyet"):
                rec["can_bo_giai_quyet"] = "; ".join(filter(None, [clean_person_name(x) for x in str(rec["can_bo_giai_quyet"]).split(";")]))
            if rec.get("can_bo_thu_ly"):
                rec["can_bo_thu_ly"] = "; ".join(filter(None, [clean_person_name(x) for x in str(rec["can_bo_thu_ly"]).split(";")]))

    wb_out = openpyxl.Workbook()
    # Sheet 1, Sheet 2, Sheet 3: Báo cáo Tổng quan Pivot Table & Rủi ro Thời hiệu phục vụ Lãnh đạo
    create_executive_pivot_sheet(wb_out, records, title_prefix="BÁO CÁO TỔNG HỢP MA TRẬN PIVOT ĐỐI CHIẾU DỮ LIỆU QLAHS", title="Tong_Quan_Pivot_Lanh_Dao", index=0)
    create_simple_stage_pivot_sheet(wb_out, records, title="Pivot_Giai_Doan_VKS", index=1)
    create_case_risk_statute_pivot_sheet(wb_out, records, title="Ma_Tran_Rui_Ro_Thoi_Hieu", index=2)

    # Sheet 4: Danh sách hợp nhất
    ws1 = wb_out.create_sheet(title="Danh_Sach_Hop_Nhat_Chuan_Hoa", index=3)
    if "Sheet" in wb_out.sheetnames:
        wb_out.remove(wb_out["Sheet"])
    ws1.views.sheetView[0].showGridLines = True
    
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )
    double_bottom_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="double", color="0F172A")
    )
    
    # 29 Columns (với Lãnh đạo Viện & Lãnh đạo Phòng tách riêng)
    COLUMNS_29 = [
        ("stt", "STT", 6, "center", "0F2942"),
        ("nguon_du_lieu", "Nguồn dữ liệu đối chiếu", 24, "center", "0F2942"),
        ("ma_vu_an", "Mã vụ án", 28, "center", "0F2942"),
        ("ten_vu_an", "Tên vụ án / Nội dung / QĐ TĐC", 38, "left", "0F2942"),
        ("tinh_thanh", "Tỉnh / Thành phố", 18, "center", "0F2942"),
        
        ("so_qd_khoi_to", "Số QĐ khởi tố", 14, "center", "1E3A8A"),
        ("ngay_qd_khoi_to", "Ngày ra QĐ khởi tố", 16, "center", "1E3A8A"),
        ("don_vi_khoi_to", "Đơn vị ra QĐ khởi tố", 30, "left", "1E3A8A"),
        
        ("so_thu_ly", "Số thụ lý", 14, "center", "0E7490"),
        ("ngay_thu_ly", "Ngày thụ lý", 16, "center", "0E7490"),
        ("loai_thu_ly", "Loại thụ lý", 16, "center", "0E7490"),
        ("vks_thu_ly", "Đơn vị VKS thụ lý", 34, "left", "0E7490"),
        
        ("ngay_tiep_nhan", "Ngày tiếp nhận hồ sơ", 16, "center", "991B1B"),
        ("thoi_han_giai_quyet", "Thời hạn giải quyết / Xử lý", 16, "center", "991B1B"),
        ("ngay_het_thoi_hieu", "Ngày hết thời hiệu TNHS", 18, "center", "991B1B"),
        ("so_ngay_con_lai", "Số ngày còn lại", 15, "center", "991B1B"),
        ("muc_do_khan_cap", "Cảnh báo tiến độ / Khẩn cấp", 28, "center", "991B1B"),
        ("canh_bao_thoi_hieu", "Cảnh báo chi tiết", 28, "center", "991B1B"),
        
        ("bi_can_dau_vu", "Bị can đầu vụ & CCCD", 26, "left", "166534"),
        ("tong_bi_can", "Tổng số bị can", 14, "center", "166534"),
        ("dieu_luat", "Thông tin tội danh / Điều luật", 36, "left", "166534"),
        ("loai_toi", "Phân loại tội", 18, "center", "166534"),
        
        ("lanh_dao_vien", "Lãnh đạo Viện phụ trách", 24, "center", "581C87"),
        ("lanh_dao_phong", "Lãnh đạo Phòng phụ trách", 28, "center", "581C87"),
        ("can_bo_giai_quyet", "Kiểm sát viên giải quyết", 28, "left", "581C87"),
        ("can_bo_thu_ly", "Cán bộ thụ lý / Chuyên viên", 22, "left", "581C87"),

        ("giai_doan", "Giai đoạn tố tụng", 16, "center", "334155"),
        ("trang_thai_vu_an", "Trạng thái vụ án", 24, "center", "334155"),
        ("trang_thai_bieu_mau", "Trạng thái biểu mẫu", 22, "left", "334155"),
        ("ten_don_vi_moi_nhat", "TÊN ĐƠN VỊ CẬP NHẬT MỚI NHẤT", 32, "left", "0E7490"),
    ]
    
    # Rows 1-3 Banner Title
    ws1.merge_cells("A1:AD1")
    ws1["A1"] = "VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH"
    ws1["A1"].font = Font(name="Segoe UI", size=12, bold=True, color="FBBF24")
    ws1["A1"].fill = PatternFill(start_color="07172C", end_color="07172C", fill_type="solid")
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 28
    
    ws1.merge_cells("A2:AD2")
    ws1["A2"] = "KẾT QUẢ RÀ SOÁT, ĐỐI CHIẾU & HỢP NHẤT DỮ LIỆU ĐA NGUỒN QLAHS"
    ws1["A2"].font = Font(name="Segoe UI", size=15, bold=True, color="FFFFFF")
    ws1["A2"].fill = PatternFill(start_color="0B2545", end_color="0B2545", fill_type="solid")
    ws1["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[2].height = 34
    
    ws1.merge_cells("A3:AD3")
    ws1["A3"] = f"★ Đã đối chiếu & hợp nhất thành công: {len(records):,} vụ án - Tách riêng Lãnh đạo Viện, Lãnh đạo Phòng & Kiểm sát viên ★"
    ws1["A3"].font = Font(name="Segoe UI", size=10, bold=True, color="047857")
    ws1["A3"].fill = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
    ws1["A3"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[3].height = 24
    
    # Row 4: Group Header Banners (7 Nhóm màu)
    GROUP_BANNERS = [
        ("A4:E4", "1. THÔNG TIN VỤ ÁN & NGUỒN ĐỐI CHIẾU", "0F2942", "F8FAFC"),
        ("F4:H4", "2. QUYẾT ĐỊNH KHỞI TỐ", "1E3A8A", "DBEAFE"),
        ("I4:L4", "3. THỤ LÝ KIỂM SÁT", "0E7490", "CFFAFE"),
        ("M4:R4", "4. TIẾN ĐỘ, QUÁ HẠN & THỜI HIỆU TNHS (ĐẶC BIỆT CHÚ Ý)", "991B1B", "FEE2E2"),
        ("S4:V4", "5. BỊ CAN & TỘI DANH", "166534", "DCFCE7"),
        ("W4:Z4", "6. PHÂN CÔNG LÃNH ĐẠO & CÁN BỘ NGHIỆP VỤ", "581C87", "F3E8FF"),
        ("AA4:AD4", "7. TIẾN ĐỘ, TRẠNG THÁI & ĐƠN VỊ MỚI NHẤT", "334155", "F1F5F9"),
    ]
    ws1.row_dimensions[4].height = 24
    for range_str, g_title, bg_color, text_color in GROUP_BANNERS:
        ws1.merge_cells(range_str)
        top_cell = ws1[range_str.split(":")[0]]
        top_cell.value = g_title
        top_cell.font = Font(name="Segoe UI", size=10, bold=True, color=text_color)
        top_cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
        top_cell.alignment = Alignment(horizontal="center", vertical="center")
        
    # Row 5: Column Headers
    ws1.row_dimensions[5].height = 28
    for col_idx, (col_key, col_title, width, align, bg_c) in enumerate(COLUMNS_29, 1):
        cell = ws1.cell(row=5, column=col_idx)
        cell.value = col_title
        cell.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=bg_c, end_color=bg_c, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        col_letter = get_column_letter(col_idx)
        ws1.column_dimensions[col_letter].width = width

    # Write Data Rows
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    
    fill_red = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    font_red = Font(name="Segoe UI", size=9, bold=True, color="991B1B")
    
    fill_amber = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    font_amber = Font(name="Segoe UI", size=9, bold=True, color="92400E")
    
    fill_yellow = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")
    font_yellow = Font(name="Segoe UI", size=9, bold=True, color="854D0E")

    fill_purple_ld = PatternFill(start_color="FAF5FF", end_color="FAF5FF", fill_type="solid")
    font_ld_v = Font(name="Segoe UI", size=9, bold=True, color="581C87")
    font_ld_p = Font(name="Segoe UI", size=9, bold=True, color="6B21A8")

    DATE_FIELDS = {"ngay_qd_khoi_to", "ngay_thu_ly", "ngay_tiep_nhan", "thoi_han_giai_quyet", "ngay_het_thoi_hieu"}

    for r_idx, rec in enumerate(records, 6):
        ws1.row_dimensions[r_idx].height = 23
        row_fill = fill_zebra if r_idx % 2 == 0 else fill_white

        # Tách và chuẩn hóa sạch tên cán bộ/lãnh đạo (loại bỏ mọi tiền tố chức danh PVT, TP, KSV...)
        if not rec.get("lanh_dao_vien") and not rec.get("lanh_dao_phong"):
            ld_v, ld_p, ksv, cb_tl = split_officers(rec.get("can_bo_giai_quyet", ""), rec.get("can_bo_thu_ly", ""))
            rec["lanh_dao_vien"] = ld_v
            rec["lanh_dao_phong"] = ld_p
            rec["can_bo_giai_quyet"] = ksv
            rec["can_bo_thu_ly"] = cb_tl
        else:
            if rec.get("lanh_dao_vien"):
                rec["lanh_dao_vien"] = "; ".join(filter(None, [clean_person_name(x) for x in str(rec["lanh_dao_vien"]).split(";")]))
            if rec.get("lanh_dao_phong"):
                rec["lanh_dao_phong"] = "; ".join(filter(None, [clean_person_name(x) for x in str(rec["lanh_dao_phong"]).split(";")]))
            if rec.get("can_bo_giai_quyet"):
                rec["can_bo_giai_quyet"] = "; ".join(filter(None, [clean_person_name(x) for x in str(rec["can_bo_giai_quyet"]).split(";")]))
            if rec.get("can_bo_thu_ly"):
                rec["can_bo_thu_ly"] = "; ".join(filter(None, [clean_person_name(x) for x in str(rec["can_bo_thu_ly"]).split(";")]))
            
        for c_idx, (col_key, _, _, align, _) in enumerate(COLUMNS_29, 1):
            cell = ws1.cell(row=r_idx, column=c_idx)
            val = rec.get(col_key, "")
            
            if col_key == "stt":
                cell.value = r_idx - 5
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_key in DATE_FIELDS:
                dt_obj = parse_date_safely(val)
                if dt_obj:
                    cell.value = dt_obj
                    cell.number_format = "DD/MM/YYYY"
                else:
                    cell.value = str(val or "")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_key == "so_ngay_con_lai":
                if val is not None and str(val).strip() != "":
                    try:
                        cell.value = int(val)
                    except ValueError:
                        cell.value = str(val)
                else:
                    cell.value = ""
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.value = str(val or "")
                cell.alignment = Alignment(horizontal=align, vertical="center")
                
            cell.border = thin_border
            cell.fill = row_fill
            cell.font = Font(name="Segoe UI", size=9)

            # Cột Lãnh đạo Viện & Lãnh đạo Phòng: Nổi bật trang trọng
            if col_key == "lanh_dao_vien" and val:
                cell.fill = fill_purple_ld
                cell.font = font_ld_v
            elif col_key == "lanh_dao_phong" and val:
                cell.fill = fill_purple_ld
                cell.font = font_ld_p

            # Cột Cảnh báo tiến độ / Khẩn cấp: Tô màu chuẩn nghiệp vụ
            if col_key in ("muc_do_khan_cap", "canh_bao_thoi_hieu") and val:
                val_str = str(val)
                val_low = val_str.lower()
                if "🔴" in val_str or "quá hạn" in val_low or "quá thời hiệu" in val_low:
                    cell.fill = fill_red
                    cell.font = font_red
                elif "🟠" in val_str:
                    cell.fill = fill_amber
                    cell.font = font_amber
                elif "🟡" in val_str:
                    cell.fill = fill_yellow
                    cell.font = font_yellow
                elif "⏸" in val_str or "tạm đình chỉ" in val_low:
                    cell.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
                    cell.font = Font(name="Segoe UI", size=9, bold=True, color="475569")
                elif "⏹" in val_str or "đình chỉ" in val_low:
                    cell.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
                    cell.font = Font(name="Segoe UI", size=9, bold=True, color="64748B")
                elif "✔" in val_str or "hoàn thành" in val_low or "đã giải quyết" in val_low:
                    cell.fill = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
                    cell.font = Font(name="Segoe UI", size=9, bold=True, color="047857")
                elif "⚖" in val_str or "đã xét xử" in val_low:
                    cell.fill = PatternFill(start_color="F5F3FF", end_color="F5F3FF", fill_type="solid")
                    cell.font = Font(name="Segoe UI", size=9, bold=True, color="581C87")
                    
    ws1.auto_filter.ref = f"A5:AD{len(records) + 5}"
    ws1.freeze_panes = "D6"

    # ==========================================
    # SHEET 5: BÁO CÁO PHÂN CÔNG & ĐỐI CHIẾU
    # ==========================================
    ws2 = wb_out.create_sheet(title="Bao_Cao_Ra_Soat_Doi_Chieu", index=4)
    ws2.views.sheetView[0].showGridLines = True
    
    ws2.merge_cells("A1:G1")
    ws2["A1"] = "BÁO CÁO TỔNG KẾT RÀ SOÁT, ĐỐI CHIẾU & PHÂN CÔNG LÃNH ĐẠO"
    ws2["A1"].font = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
    ws2["A1"].fill = PatternFill(start_color="0B2545", end_color="0B2545", fill_type="solid")
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 32

    # Trích xuất danh sách nguồn
    multi_source_count = sum(1 for r in records if "Khớp nối" in str(r.get("nguon_du_lieu", "")))
    all_source_names = set()
    for r in records:
        src = str(r.get("nguon_du_lieu", ""))
        if "Khớp nối" in src and ":" in src:
            parts = src.split(":", 1)[1].split("+")
            for p in parts:
                all_source_names.add(p.strip())
        elif src:
            all_source_names.add(src.strip())
    n_source_files = max(1, len(all_source_names))

    # Bảng 1: Chỉ số tổng thể
    metrics = [
        ("Tổng số vụ án sau khi rà soát & hợp nhất", len(records), "Tổng hợp trọn vẹn từ tất cả các file kết xuất QLAHS"),
        ("Tổng số tập tin nguồn QLAHS đã nạp & đối chiếu", n_source_files, f"Bao gồm: {', '.join(sorted(all_source_names))[:80]}..."),
        ("Số vụ án được khớp nối & bù đắp từ nhiều file", multi_source_count, "Được tự động đắp bù số thụ lý, KSV, hạn điều tra, thời hiệu TNHS"),
        ("Số vụ án có cảnh báo quá hạn thời hạn điều tra", sum(1 for r in records if "quá hạn" in str(r.get("canh_bao_thoi_hieu", "")).lower()), "Cần lãnh đạo VKS và KSV chỉ đạo điều tra giải quyết gấp"),
        ("Số vụ án đã phân công Lãnh đạo Viện phụ trách", sum(1 for r in records if r.get("lanh_dao_vien")), "Ghi nhận đích danh Lãnh đạo Viện phụ trách chỉ đạo"),
        ("Số vụ án đã phân công Lãnh đạo Phòng phụ trách", sum(1 for r in records if r.get("lanh_dao_phong")), "Ghi nhận đích danh Lãnh đạo Phòng 2 phụ trách"),
        ("Số vụ án đã phân công Kiểm sát viên giải quyết", sum(1 for r in records if r.get("can_bo_giai_quyet")), "Có tên KSV thực hành quyền công tố & kiểm sát điều tra"),
        ("Số vụ án có thông tin tội danh & điều luật đầy đủ", sum(1 for r in records if r.get("dieu_luat")), "Trích xuất chi tiết từng điều khoản BLHS"),
    ]
    
    headers_s2 = ["STT", "Chỉ số rà soát & đối chiếu", "Số lượng", "Đơn vị tính", "Đánh giá & Khuyến nghị nghiệp vụ"]
    ws2.row_dimensions[3].height = 24
    for c_idx, h in enumerate(headers_s2, 1):
        cell = ws2.cell(row=3, column=c_idx, value=h)
        cell.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="0F2942", end_color="0F2942", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
    for m_idx, (m_title, m_val, m_note) in enumerate(metrics, 4):
        ws2.row_dimensions[m_idx].height = 22
        ws2.cell(row=m_idx, column=1, value=m_idx - 3).alignment = Alignment(horizontal="center", vertical="center")
        ws2.cell(row=m_idx, column=2, value=m_title).alignment = Alignment(horizontal="left", vertical="center")
        c3 = ws2.cell(row=m_idx, column=3, value=m_val)
        c3.alignment = Alignment(horizontal="center", vertical="center")
        c3.font = Font(name="Segoe UI", size=10, bold=True, color="1E3A8A")
        ws2.cell(row=m_idx, column=4, value="Vụ án").alignment = Alignment(horizontal="center", vertical="center")
        ws2.cell(row=m_idx, column=5, value=m_note).alignment = Alignment(horizontal="left", vertical="center")
        for c in range(1, 6):
            ws2.cell(row=m_idx, column=c).border = thin_border

    # Bảng 2: Thống kê chi tiết phân công theo Lãnh đạo Viện & Lãnh đạo Phòng
    start_r2 = len(metrics) + 6
    ws2.merge_cells(f"A{start_r2}:F{start_r2}")
    t2_cell = ws2[f"A{start_r2}"]
    t2_cell.value = "THỐNG KÊ CHI TIẾT SỐ LƯỢNG VỤ ÁN THEO LÃNH ĐẠO VIỆN & LÃNH ĐẠO PHÒNG 2"
    t2_cell.font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    t2_cell.fill = PatternFill(start_color="581C87", end_color="581C87", fill_type="solid")
    t2_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[start_r2].height = 26

    # Đếm số vụ án của từng lãnh đạo (Chuẩn hóa hợp nhất không trùng lặp dòng)
    leader_stats = []
    
    # 1. Lãnh đạo Viện (Hợp nhất toàn bộ các biến thể về tên chính thức, không hiển thị tiền tố PVT)
    for name in LANH_DAO_VIEN:
        cnt = sum(1 for r in records if name.lower() in str(r.get("lanh_dao_vien", "")).lower())
        pct = (cnt / len(records) * 100) if records else 0
        leader_stats.append(("Lãnh đạo Viện", name, "Phó Viện trưởng VKSND TP.HCM", cnt, pct))

    # Phó Viện trưởng khác nếu có (Trần Văn Cường, Bảo, Tùng...) - chỉ lấy họ tên sạch
    other_pvts = set()
    for r in records:
        for p in str(r.get("lanh_dao_vien", "")).split(";"):
            p_clean = clean_person_name(p.strip())
            if p_clean and p_clean not in LANH_DAO_VIEN:
                other_pvts.add(p_clean)
    for p in sorted(other_pvts):
        cnt = sum(1 for r in records if p.lower() in str(r.get("lanh_dao_vien", "")).lower())
        pct = (cnt / len(records) * 100) if records else 0
        leader_stats.append(("Lãnh đạo Viện", p, "Phó Viện trưởng phụ trách", cnt, pct))

    # 2. Lãnh đạo Phòng 2 (Họ và tên chuẩn không kèm tiền tố TP/PTP)
    for idx_p, name in enumerate(LANH_DAO_PHONG_2):
        role = "Trưởng phòng 2" if idx_p == 0 else "Phó Trưởng phòng 2"
        cnt = sum(1 for r in records if name.lower() in str(r.get("lanh_dao_phong", "")).lower())
        pct = (cnt / len(records) * 100) if records else 0
        leader_stats.append(("Lãnh đạo Phòng 2", name, role, cnt, pct))

    # Lãnh đạo Phòng khác nếu có
    other_ptps = set()
    for r in records:
        for p in str(r.get("lanh_dao_phong", "")).split(";"):
            p_clean = clean_person_name(p.strip())
            if p_clean and p_clean not in LANH_DAO_PHONG_2:
                other_ptps.add(p_clean)
    for p in sorted(other_ptps):
        cnt = sum(1 for r in records if p.lower() in str(r.get("lanh_dao_phong", "")).lower())
        pct = (cnt / len(records) * 100) if records else 0
        leader_stats.append(("Lãnh đạo Phòng 2", p, "Phó Trưởng phòng phụ trách", cnt, pct))

    headers_ld = ["STT", "Cấp Lãnh đạo", "Họ và tên Lãnh đạo", "Chức vụ / Phụ trách", "Số lượng vụ án", "Tỷ lệ %"]
    ws2.row_dimensions[start_r2 + 1].height = 24
    for c_idx, h in enumerate(headers_ld, 1):
        cell = ws2.cell(row=start_r2 + 1, column=c_idx, value=h)
        cell.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="3B0764", end_color="3B0764", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    for l_idx, (level, name, role, cnt, pct) in enumerate(leader_stats, start_r2 + 2):
        ws2.row_dimensions[l_idx].height = 22
        ws2.cell(row=l_idx, column=1, value=l_idx - start_r2 - 1).alignment = Alignment(horizontal="center", vertical="center")
        
        c_lev = ws2.cell(row=l_idx, column=2, value=level)
        c_lev.alignment = Alignment(horizontal="center", vertical="center")
        if "Viện" in level:
            c_lev.font = Font(name="Segoe UI", size=9, bold=True, color="581C87")
        else:
            c_lev.font = Font(name="Segoe UI", size=9, bold=True, color="0E7490")
            
        c_name = ws2.cell(row=l_idx, column=3, value=name)
        c_name.alignment = Alignment(horizontal="left", vertical="center")
        c_name.font = Font(name="Segoe UI", size=9, bold=True)
        
        ws2.cell(row=l_idx, column=4, value=role).alignment = Alignment(horizontal="left", vertical="center")
        
        c_cnt = ws2.cell(row=l_idx, column=5, value=cnt)
        c_cnt.alignment = Alignment(horizontal="center", vertical="center")
        c_cnt.font = Font(name="Segoe UI", size=10, bold=True, color="991B1B" if cnt > 0 else "64748B")
        
        c_pct = ws2.cell(row=l_idx, column=6, value=f"{pct:.1f}%")
        c_pct.alignment = Alignment(horizontal="center", vertical="center")
        
        for c in range(1, 7):
            ws2.cell(row=l_idx, column=c).border = thin_border

    ws2.column_dimensions["A"].width = 8
    ws2.column_dimensions["B"].width = 22
    ws2.column_dimensions["C"].width = 26
    ws2.column_dimensions["D"].width = 30
    ws2.column_dimensions["E"].width = 18
    ws2.column_dimensions["F"].width = 14

    # ==========================================
    # SHEET 6: PIVOT TABLE TỔNG HỢP VKS x LOẠI TỘI
    # ==========================================
    ws3 = wb_out.create_sheet(title="Tong_Hop_Pivot_Table", index=5)
    ws3.views.sheetView[0].showGridLines = True
    
    ws3.merge_cells("A1:G1")
    ws3["A1"] = "BẢNG TỔNG HỢP PIVOT TABLE THEO ĐƠN VỊ VKS VÀ PHÂN LOẠI TỘI PHẠM"
    ws3["A1"].font = Font(name="Segoe UI", size=12, bold=True, color="FFFFFF")
    ws3["A1"].fill = PatternFill(start_color="0B2545", end_color="0B2545", fill_type="solid")
    ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 32

    # Tổng hợp ma trận
    unit_stats = {}
    for r in records:
        u = r.get("vks_thu_ly") or "Chưa rõ VKS"
        lt = (r.get("loai_toi") or "").lower()
        if u not in unit_stats:
            unit_stats[u] = {"total": 0, "it": 0, "nt": 0, "rnt": 0, "dbnt": 0}
        unit_stats[u]["total"] += 1
        if "ít" in lt:
            unit_stats[u]["it"] += 1
        elif "rất" in lt:
            unit_stats[u]["rnt"] += 1
        elif "đặc biệt" in lt:
            unit_stats[u]["dbnt"] += 1
        elif "nghiêm trọng" in lt:
            unit_stats[u]["nt"] += 1

    headers_s3 = ["STT", "Đơn vị VKS thụ lý", "Tổng số vụ", "Ít nghiêm trọng", "Nghiêm trọng", "Rất nghiêm trọng", "Đặc biệt nghiêm trọng"]
    ws3.row_dimensions[3].height = 24
    for c_idx, h in enumerate(headers_s3, 1):
        cell = ws3.cell(row=3, column=c_idx, value=h)
        cell.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="0F2942", end_color="0F2942", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    sorted_units = sorted(unit_stats.keys(), key=lambda x: unit_stats[x]["total"], reverse=True)
    tot_all = len(records)
    tot_it = sum(v["it"] for v in unit_stats.values())
    tot_nt = sum(v["nt"] for v in unit_stats.values())
    tot_rnt = sum(v["rnt"] for v in unit_stats.values())
    tot_dbnt = sum(v["dbnt"] for v in unit_stats.values())

    for u_idx, u_name in enumerate(sorted_units, 4):
        ws3.row_dimensions[u_idx].height = 22
        st = unit_stats[u_name]
        ws3.cell(row=u_idx, column=1, value=u_idx - 3).alignment = Alignment(horizontal="center", vertical="center")
        ws3.cell(row=u_idx, column=2, value=u_name).alignment = Alignment(horizontal="left", vertical="center")
        ws3.cell(row=u_idx, column=3, value=st["total"]).alignment = Alignment(horizontal="center", vertical="center")
        ws3.cell(row=u_idx, column=4, value=st["it"]).alignment = Alignment(horizontal="center", vertical="center")
        ws3.cell(row=u_idx, column=5, value=st["nt"]).alignment = Alignment(horizontal="center", vertical="center")
        ws3.cell(row=u_idx, column=6, value=st["rnt"]).alignment = Alignment(horizontal="center", vertical="center")
        ws3.cell(row=u_idx, column=7, value=st["dbnt"]).alignment = Alignment(horizontal="center", vertical="center")
        for c in range(1, 8):
            ws3.cell(row=u_idx, column=c).border = thin_border

    # Row Tổng cộng
    tot_row = len(sorted_units) + 4
    ws3.row_dimensions[tot_row].height = 24
    ws3.cell(row=tot_row, column=1, value="∑").alignment = Alignment(horizontal="center", vertical="center")
    ws3.cell(row=tot_row, column=2, value="TỔNG CỘNG CHUNG").alignment = Alignment(horizontal="center", vertical="center")
    ws3.cell(row=tot_row, column=3, value=tot_all).alignment = Alignment(horizontal="center", vertical="center")
    ws3.cell(row=tot_row, column=4, value=tot_it).alignment = Alignment(horizontal="center", vertical="center")
    ws3.cell(row=tot_row, column=5, value=tot_nt).alignment = Alignment(horizontal="center", vertical="center")
    ws3.cell(row=tot_row, column=6, value=tot_rnt).alignment = Alignment(horizontal="center", vertical="center")
    ws3.cell(row=tot_row, column=7, value=tot_dbnt).alignment = Alignment(horizontal="center", vertical="center")
    
    for c in range(1, 8):
        cell = ws3.cell(row=tot_row, column=c)
        cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        cell.border = double_bottom_border

    ws3.column_dimensions["A"].width = 8
    ws3.column_dimensions["B"].width = 52
    ws3.column_dimensions["C"].width = 16
    ws3.column_dimensions["D"].width = 18
    ws3.column_dimensions["E"].width = 18
    ws3.column_dimensions["F"].width = 20
    ws3.column_dimensions["G"].width = 24

    wb_out.active = 0  # Sheet 1: Tong_Quan_Pivot_Lanh_Dao
    wb_out.save(out_filepath)


def export_case_enriched_to_excel(case_records, linked_defs, unlinked_defs, out_filepath):
    """
    Xuất file Excel liên thông & làm giàu dữ liệu Sổ vụ án ⟷ Bị can QLAHS.
    Gồm 5 Sheet chuyên nghiệp:
    - Sheet 1: Tong_Quan_Pivot_Lanh_Dao (Báo cáo Đa Chiều Toàn Diện Lãnh Đạo)
    - Sheet 2: Pivot_Giai_Doan_VKS (Ma Trận Pivot Đơn vị VKS x Giai đoạn tố tụng)
    - Sheet 3: So_Vu_An_Da_Lam_Giau (Sổ vụ án 24 cột đã bổ sung chi tiết danh sách bị can)
    - Sheet 4: Chi_Tiet_Bi_Can_Lien_Thong (15 cột bị can đã khớp chính xác với hồ sơ vụ án)
    - Sheet 5: Bi_Can_Cho_Xac_Minh_PM_Cu (12 cột bị can mã cũ chờ đối soát từ phần mềm cũ)
    """
    for rec in case_records:
        normalize_case_stage_and_status(rec)
        if not rec.get("lanh_dao_vien") and not rec.get("lanh_dao_phong"):
            ld_v, ld_p, ksv, cb_tl = split_officers(rec.get("can_bo_giai_quyet", ""), rec.get("can_bo_thu_ly", ""))
            rec["lanh_dao_vien"] = ld_v
            rec["lanh_dao_phong"] = ld_p
            rec["can_bo_giai_quyet"] = ksv
            rec["can_bo_thu_ly"] = cb_tl
        else:
            if rec.get("lanh_dao_vien"):
                rec["lanh_dao_vien"] = "; ".join(filter(None, [clean_person_name(x) for x in str(rec["lanh_dao_vien"]).split(";")]))
            if rec.get("lanh_dao_phong"):
                rec["lanh_dao_phong"] = "; ".join(filter(None, [clean_person_name(x) for x in str(rec["lanh_dao_phong"]).split(";")]))
            if rec.get("can_bo_giai_quyet"):
                rec["can_bo_giai_quyet"] = "; ".join(filter(None, [clean_person_name(x) for x in str(rec["can_bo_giai_quyet"]).split(";")]))
            if rec.get("can_bo_thu_ly"):
                rec["can_bo_thu_ly"] = "; ".join(filter(None, [clean_person_name(x) for x in str(rec["can_bo_thu_ly"]).split(";")]))

    wb_out = openpyxl.Workbook()
    create_executive_pivot_sheet(wb_out, case_records, title_prefix="BÁO CÁO TỔNG HỢP MA TRẬN PIVOT VỤ ÁN ĐÃ LIÊN THÔNG BỊ CAN", title="Tong_Quan_Pivot_Lanh_Dao", index=0)
    create_simple_stage_pivot_sheet(wb_out, case_records, title="Pivot_Giai_Doan_VKS", index=1)

    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )

    # Sheet 3: Sổ vụ án đã làm giàu
    ws3 = wb_out.create_sheet(title="So_Vu_An_Da_Lam_Giau", index=2)
    if "Sheet" in wb_out.sheetnames:
        wb_out.remove(wb_out["Sheet"])
    ws3.views.sheetView[0].showGridLines = True

    COLUMNS_CASE = [
        ("stt", "STT", 6, "center", "0F2942"),
        ("nguon_du_lieu", "Nguồn dữ liệu đối chiếu", 24, "center", "0F2942"),
        ("ma_vu_an", "Mã vụ án", 28, "center", "0F2942"),
        ("ten_vu_an", "Tên vụ án / Nội dung", 38, "left", "0F2942"),
        ("tinh_thanh", "Tỉnh / Thành phố", 18, "center", "0F2942"),
        ("so_qd_khoi_to", "Số QĐ khởi tố", 14, "center", "1E3A8A"),
        ("ngay_qd_khoi_to", "Ngày ra QĐ khởi tố", 16, "center", "1E3A8A"),
        ("don_vi_khoi_to", "Đơn vị ra QĐ", 28, "left", "1E3A8A"),
        ("so_thu_ly", "Số thụ lý", 14, "center", "0E7490"),
        ("ngay_thu_ly", "Ngày thụ lý", 16, "center", "0E7490"),
        ("loai_thu_ly", "Loại thụ lý", 16, "center", "0E7490"),
        ("vks_thu_ly", "Đơn vị VKS thụ lý", 34, "left", "0E7490"),
        ("thoi_han_giai_quyet", "Thời hạn giải quyết", 16, "center", "991B1B"),
        ("bi_can_dau_vu", "Bị can đầu vụ & CCCD", 28, "left", "166534"),
        ("tong_bi_can", "Tổng số bị can", 14, "center", "166534"),
        ("danh_sach_bi_can_chi_tiet", "Danh sách chi tiết bị can đã làm giàu", 46, "left", "166534"),
        ("dieu_luat", "Tội danh / Điều luật", 36, "left", "166534"),
        ("loai_toi", "Phân loại tội", 18, "center", "166534"),
        ("lanh_dao_vien", "Lãnh đạo Viện", 22, "center", "581C87"),
        ("lanh_dao_phong", "Lãnh đạo Phòng", 24, "center", "581C87"),
        ("can_bo_giai_quyet", "Kiểm sát viên", 24, "left", "581C87"),
        ("can_bo_thu_ly", "Cán bộ thụ lý", 20, "left", "581C87"),
        ("giai_doan", "Giai đoạn tố tụng", 16, "center", "334155"),
        ("trang_thai_vu_an", "Trạng thái vụ án", 24, "center", "334155"),
    ]

    # Banner rows 1-3
    ws3.merge_cells("A1:X1")
    ws3["A1"] = "VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH"
    ws3["A1"].font = Font(name="Segoe UI", size=12, bold=True, color="FBBF24")
    ws3["A1"].fill = PatternFill(start_color="07172C", end_color="07172C", fill_type="solid")
    ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 28

    ws3.merge_cells("A2:X2")
    ws3["A2"] = "SỔ QUẢN LÝ VỤ ÁN HÌNH SỰ ĐÃ LIÊN THÔNG & LÀM GIÀU DANH SÁCH BỊ CAN"
    ws3["A2"].font = Font(name="Segoe UI", size=15, bold=True, color="FFFFFF")
    ws3["A2"].fill = PatternFill(start_color="0B2545", end_color="0B2545", fill_type="solid")
    ws3["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[2].height = 34

    ws3.merge_cells("A3:X3")
    enriched_cnt = sum(1 for c in case_records if c.get("danh_sach_bi_can_chi_tiet"))
    ws3["A3"] = f"★ Tổng số: {len(case_records):,} vụ án — Đã liên thông & làm giàu thành công: {enriched_cnt:,} vụ án ({len(linked_defs):,} bị can) ★"
    ws3["A3"].font = Font(name="Segoe UI", size=10, bold=True, color="047857")
    ws3["A3"].fill = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
    ws3["A3"].alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[3].height = 24

    ws3.row_dimensions[4].height = 28
    for col_idx, (fkey, label, width, align_h, col_hex) in enumerate(COLUMNS_CASE, 1):
        cell = ws3.cell(4, col_idx, label)
        cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=col_hex, end_color=col_hex, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        ws3.column_dimensions[get_column_letter(col_idx)].width = width

    for r_idx, c in enumerate(case_records, 5):
        ws3.row_dimensions[r_idx].height = 24
        for c_idx, (k, _, _, align_h, _) in enumerate(COLUMNS_CASE, 1):
            val = c.get(k, "")
            cell = ws3.cell(r_idx, c_idx, val)
            cell.font = Font(name="Segoe UI", size=10)
            cell.border = thin_border
            if k in ("stt", "tong_bi_can", "so_qd_khoi_to", "so_thu_ly", "ngay_qd_khoi_to", "ngay_thu_ly", "thoi_han_giai_quyet", "giai_doan", "trang_thai_vu_an"):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif k in ("ma_vu_an",):
                cell.number_format = "@"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            if c.get("danh_sach_bi_can_chi_tiet") and k == "danh_sach_bi_can_chi_tiet":
                cell.fill = PatternFill(start_color="F0FDF4", end_color="F0FDF4", fill_type="solid")

    ws3.freeze_panes = "A5"
    ws3.auto_filter.ref = f"A4:{get_column_letter(len(COLUMNS_CASE))}{len(case_records)+4}"

    # Sheet 4: Chi tiết bị can đã liên thông
    ws4 = wb_out.create_sheet(title="Chi_Tiet_Bi_Can_Lien_Thong", index=3)
    ws4.views.sheetView[0].showGridLines = True

    COLUMNS_DEFS = [
        ("stt", "STT", 6),
        ("ma_vu_an_linked", "Mã vụ án liên thông", 28),
        ("ten_vu_an_linked", "Tên vụ án / Nội dung", 36),
        ("ma_bi_can", "Mã bị can", 28),
        ("ten_bi_can", "Họ và tên bị can", 22),
        ("ngay_sinh", "Ngày sinh", 14),
        ("gioi_tinh", "Giới tính", 10),
        ("so_cccd", "Số CCCD / CMND", 18),
        ("dia_chi", "Địa chỉ cư trú", 35),
        ("dieu_luat", "Tội danh / Điều luật khởi tố", 35),
        ("loai_toi", "Loại tội", 18),
        ("trang_thai", "Trạng thái bị can", 20),
        ("giai_doan_linked", "Giai đoạn vụ án", 16),
        ("trang_thai_va_linked", "Trạng thái vụ án", 22),
        ("vks_thu_ly_linked", "Đơn vị VKS thụ lý", 32),
    ]

    ws4.merge_cells("A1:O1")
    ws4["A1"] = "DANH SÁCH CHI TIẾT BỊ CAN ĐÃ LIÊN THÔNG CHÍNH XÁC VỚI HỒ SƠ VỤ ÁN"
    ws4["A1"].font = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
    ws4["A1"].fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    ws4["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws4.row_dimensions[1].height = 30

    ws4.row_dimensions[2].height = 26
    for c_idx, (_, title, w) in enumerate(COLUMNS_DEFS, 1):
        cell = ws4.cell(2, c_idx, title)
        cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="0F2942", end_color="0F2942", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        ws4.column_dimensions[get_column_letter(c_idx)].width = w

    for r_idx, d in enumerate(linked_defs, 3):
        ws4.row_dimensions[r_idx].height = 24
        for c_idx, (k, _, _) in enumerate(COLUMNS_DEFS, 1):
            val = d.get(k, "")
            cell = ws4.cell(r_idx, c_idx, val)
            cell.font = Font(name="Segoe UI", size=10)
            cell.border = thin_border
            if k in ("stt", "ngay_sinh", "gioi_tinh", "giai_doan_linked"):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif k in ("ma_vu_an_linked", "ma_bi_can", "so_cccd"):
                cell.number_format = "@"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws4.freeze_panes = "A3"
    if linked_defs:
        ws4.auto_filter.ref = f"A2:{get_column_letter(len(COLUMNS_DEFS))}{len(linked_defs)+2}"

    # Sheet 5: Bị can chờ xác minh PM cũ
    ws5 = wb_out.create_sheet(title="Bi_Can_Cho_Xac_Minh_PM_Cu", index=4)
    ws5.views.sheetView[0].showGridLines = True

    COLUMNS_UNLINKED = [
        ("stt", "STT", 6),
        ("ma_bi_can", "Mã bị can", 28),
        ("ten_bi_can", "Họ và tên bị can", 22),
        ("ngay_sinh", "Ngày sinh", 14),
        ("gioi_tinh", "Giới tính", 10),
        ("so_cccd", "Số CCCD / CMND", 18),
        ("dia_chi", "Địa chỉ cư trú", 35),
        ("dieu_luat", "Tội danh / Điều luật khởi tố", 35),
        ("loai_toi", "Loại tội", 18),
        ("trang_thai", "Trạng thái giải quyết", 20),
        ("ma_vu_an_linked", "Tình trạng mã vụ án", 20),
        ("ten_vu_an_linked", "Ghi chú đối soát", 36),
    ]

    ws5.merge_cells("A1:L1")
    ws5["A1"] = "DANH SÁCH BỊ CAN CHỜ ĐỐI SOÁT TỪ PHẦN MỀM CŨ / BỔ SUNG SỔ VỤ ÁN"
    ws5["A1"].font = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
    ws5["A1"].fill = PatternFill(start_color="991B1B", end_color="991B1B", fill_type="solid")
    ws5["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws5.row_dimensions[1].height = 30

    ws5.row_dimensions[2].height = 26
    for c_idx, (_, title, w) in enumerate(COLUMNS_UNLINKED, 1):
        cell = ws5.cell(2, c_idx, title)
        cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="7F1D1D", end_color="7F1D1D", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        ws5.column_dimensions[get_column_letter(c_idx)].width = w

    for r_idx, d in enumerate(unlinked_defs, 3):
        ws5.row_dimensions[r_idx].height = 24
        for c_idx, (k, _, _) in enumerate(COLUMNS_UNLINKED, 1):
            val = d.get(k, "")
            cell = ws5.cell(r_idx, c_idx, val)
            cell.font = Font(name="Segoe UI", size=10)
            cell.border = thin_border
            if k == "ma_vu_an_linked":
                cell.font = Font(name="Segoe UI", size=10, bold=True, color="DC2626")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.fill = PatternFill(start_color="FEF2F2", end_color="FEF2F2", fill_type="solid")
            elif k in ("stt", "ngay_sinh", "gioi_tinh"):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif k in ("ma_bi_can", "so_cccd"):
                cell.number_format = "@"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws5.freeze_panes = "A3"
    if unlinked_defs:
        ws5.auto_filter.ref = f"A2:{get_column_letter(len(COLUMNS_UNLINKED))}{len(unlinked_defs)+2}"

    wb_out.active = 0  # Sheet 1: Tong_Quan_Pivot_Lanh_Dao
    wb_out.save(out_filepath)


def build_pivot_summary(records):
    """Tính toán ma trận Pivot Table: Tỉnh/Thành × Giai đoạn và Trạng thái hồ sơ."""
    if not records:
        return {}, [], []

    stages_order = ["Điều tra", "Truy tố", "Sơ thẩm", "Phúc thẩm", "Thi hành án"]
    found_stages = []
    for r in records:
        gd = r.get("giai_doan", "").strip() or "Chưa rõ"
        if gd not in found_stages:
            found_stages.append(gd)

    stages = [s for s in stages_order if s in found_stages]
    for s in found_stages:
        if s not in stages:
            stages.append(s)

    matrix = {}
    for r in records:
        prov = r.get("tinh_thanh", "").strip() or "Khác / Chưa rõ"
        gd = r.get("giai_doan", "").strip() or "Chưa rõ"
        tt = r.get("trang_thai", "").strip() or "Khác"
        va = r.get("ma_vu_an", "").strip()

        if prov not in matrix:
            matrix[prov] = {
                "total_defendants": 0,
                "cases": set(),
                "stages": {s: 0 for s in stages},
                "status": {},
            }
        matrix[prov]["total_defendants"] += 1
        if va:
            matrix[prov]["cases"].add(va)
        if gd not in matrix[prov]["stages"]:
            matrix[prov]["stages"][gd] = 0
        matrix[prov]["stages"][gd] += 1
        matrix[prov]["status"][tt] = matrix[prov]["status"].get(tt, 0) + 1

    sorted_provinces = sorted(matrix.keys(), key=lambda p: matrix[p]["total_defendants"], reverse=True)
    return matrix, stages, sorted_provinces


def export_records_to_excel(records, out_filepath):
    """Xuất file Excel gồm 2 Sheet: Sheet 1 Danh sách 12 cột, Sheet 2 Bảng Pivot Table Tổng hợp."""
    wb = openpyxl.Workbook()

    # Sheet 1: Danh sách chuẩn hóa 12 cột
    ws1 = wb.active
    ws1.title = "Danh_Sach_Chuan_Hoa"
    ws1.views.sheetView[0].showGridLines = True

    NAVY_VKS = "0A2540"     # Xanh Kiểm sát đậm
    CRIMSON_VKS = "991B1B"  # Đỏ Tư pháp
    GOLD_VKS = "D97706"     # Vàng Phù hiệu

    thin_border = Border(
        left=Side(style="thin", color="D1D5DB"),
        right=Side(style="thin", color="D1D5DB"),
        top=Side(style="thin", color="D1D5DB"),
        bottom=Side(style="thin", color="D1D5DB"),
    )

    headers12 = [col_title for _, col_title, _ in COLUMNS_12]
    ws1.row_dimensions[1].height = 28
    for c_idx, h_title in enumerate(headers12, start=1):
        cell = ws1.cell(row=1, column=c_idx, value=h_title)
        cell.font = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=NAVY_VKS, end_color=NAVY_VKS, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    for idx, rec in enumerate(records, start=1):
        r = idx + 1
        ws1.row_dimensions[r].height = 26
        vals = [
            int(rec["stt"]) if str(rec["stt"]).isdigit() else idx,
            rec["ma_vu_an"],
            rec["ten_vu_an"],
            rec["cqdt"],
            rec["tinh_thanh"],
            rec["giai_doan"],
            rec["ma_bi_can"],
            rec["ten_bi_can"],
            rec["ngay_sinh"],
            rec["so_cccd"],
            rec["dieu_luat"],
            rec["trang_thai"],
        ]
        for c_idx, val in enumerate(vals, start=1):
            c = ws1.cell(row=r, column=c_idx, value=val)
            c.font = Font(name="Times New Roman", size=11, bold=(c_idx in (2, 5, 8)))
            c.border = thin_border
            if c_idx == 1:
                c.alignment = Alignment(horizontal="right", vertical="center")
            elif c_idx in (2, 7, 9, 10):
                c.number_format = "@"
                c.alignment = Alignment(horizontal="left", vertical="center")
            elif c_idx in (5, 6, 12):
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws1.freeze_panes = "A2"
    if records:
        ws1.auto_filter.ref = f"A1:L{len(records)+1}"
    widths12 = [8, 24, 34, 28, 20, 14, 26, 22, 14, 16, 36, 18]
    for c_idx, w in enumerate(widths12, start=1):
        ws1.column_dimensions[get_column_letter(c_idx)].width = w

    # Sheet 2: Bảng Pivot Table Tổng hợp theo Tỉnh/Thành & Giai đoạn
    ws2 = wb.create_sheet(title="Tong_Hop_Pivot_Table")
    ws2.views.sheetView[0].showGridLines = True

    matrix, stages, sorted_provinces = build_pivot_summary(records)

    ws2.merge_cells("A1:G1")
    t_cell = ws2.cell(row=1, column=1, value="VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH")
    t_cell.font = Font(name="Times New Roman", size=13, bold=True, color="FFFFFF")
    t_cell.fill = PatternFill(start_color=NAVY_VKS, end_color=NAVY_VKS, fill_type="solid")
    t_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 30

    ws2.merge_cells("A2:G2")
    sub_cell = ws2.cell(row=2, column=1, value="HỆ THỐNG TỔNG HỢP VÀ XỬ LÝ DỮ LIỆU NTQLAHS - BÁO CÁO PIVOT TABLE TỔNG HỢP")
    sub_cell.font = Font(name="Times New Roman", size=11, bold=True, color="991B1B")
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[2].height = 24

    pivot_headers = ["STT", "Tỉnh / Thành phố", "Số vụ án", "Tổng số bị can"] + [f"GĐ: {s}" for s in stages] + ["Tỷ lệ %"]
    ws2.row_dimensions[4].height = 26
    for c_idx, h_text in enumerate(pivot_headers, start=1):
        cell = ws2.cell(row=4, column=c_idx, value=h_text)
        cell.font = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=CRIMSON_VKS, end_color=CRIMSON_VKS, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    total_records = len(records)
    current_row = 5
    for p_idx, prov in enumerate(sorted_provinces, start=1):
        data = matrix[prov]
        ws2.row_dimensions[current_row].height = 24
        pct = (data["total_defendants"] / total_records * 100) if total_records > 0 else 0
        row_vals = [
            p_idx,
            prov,
            len(data["cases"]),
            data["total_defendants"],
        ] + [data["stages"].get(s, 0) for s in stages] + [f"{pct:.1f}%"]

        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws2.cell(row=current_row, column=c_idx, value=val)
            cell.font = Font(name="Times New Roman", size=11, bold=(c_idx in (2, 4)))
            cell.border = thin_border
            if c_idx in (1, 3, 4) or c_idx > 4:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        current_row += 1

    # Dòng Tổng cộng chung
    ws2.row_dimensions[current_row].height = 26
    ws2.cell(row=current_row, column=1, value="").border = thin_border
    c_total_lbl = ws2.cell(row=current_row, column=2, value="TỔNG CỘNG CHUNG")
    c_total_lbl.font = Font(name="Times New Roman", size=11, bold=True, color=NAVY_VKS)
    c_total_lbl.border = thin_border
    c_total_lbl.alignment = Alignment(horizontal="center", vertical="center")

    all_cases = len(set(r.get("ma_vu_an", "").strip() for r in records if r.get("ma_vu_an", "").strip()))
    c_cases = ws2.cell(row=current_row, column=3, value=all_cases)
    c_cases.font = Font(name="Times New Roman", size=11, bold=True, color=NAVY_VKS)
    c_cases.border = thin_border
    c_cases.alignment = Alignment(horizontal="center", vertical="center")

    c_defs = ws2.cell(row=current_row, column=4, value=total_records)
    c_defs.font = Font(name="Times New Roman", size=11, bold=True, color=NAVY_VKS)
    c_defs.border = thin_border
    c_defs.alignment = Alignment(horizontal="center", vertical="center")

    for s_idx, s in enumerate(stages, start=5):
        s_total = sum(matrix[p]["stages"].get(s, 0) for p in sorted_provinces)
        c_stg = ws2.cell(row=current_row, column=s_idx, value=s_total)
        c_stg.font = Font(name="Times New Roman", size=11, bold=True, color=NAVY_VKS)
        c_stg.border = thin_border
        c_stg.alignment = Alignment(horizontal="center", vertical="center")

    pct_col_idx = 5 + len(stages)
    c_pct = ws2.cell(row=current_row, column=pct_col_idx, value="100.0%")
    c_pct.font = Font(name="Times New Roman", size=11, bold=True, color=NAVY_VKS)
    c_pct.border = thin_border
    c_pct.alignment = Alignment(horizontal="center", vertical="center")

    for c_idx in range(1, pct_col_idx + 1):
        cell = ws2.cell(row=current_row, column=c_idx)
        cell.fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

    ws2.column_dimensions["A"].width = 8
    ws2.column_dimensions["B"].width = 24
    ws2.column_dimensions["C"].width = 14
    ws2.column_dimensions["D"].width = 16
    for idx in range(len(stages)):
        col_letter = get_column_letter(5 + idx)
        ws2.column_dimensions[col_letter].width = 16
    ws2.column_dimensions[get_column_letter(pct_col_idx)].width = 14

    wb.active = 0
    wb.save(out_filepath)


class LoadingDialog(tk.Toplevel):
    """Hộp thoại loading chuyên nghiệp hiển thị thanh tiến trình và thông báo thời gian thực."""

    def __init__(self, parent, title="Đang xử lý dữ liệu", message="Hệ thống đang xử lý dữ liệu hồ sơ QLAHS..."):
        super().__init__(parent)
        self.title(title)
        self.transient(parent)
        self.resizable(False, False)
        self.configure(bg="#0B2545")

        # Cài đặt biểu tượng cửa sổ nếu có
        icon_path = get_resource_path("logo_vks.ico")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        # Kích thước và định vị căn giữa màn hình cha
        width, height = 480, 195
        try:
            parent.update_idletasks()
            pw = parent.winfo_width()
            ph = parent.winfo_height()
            px = parent.winfo_rootx()
            py = parent.winfo_rooty()
            x = px + max(0, (pw - width) // 2)
            y = py + max(0, (ph - height) // 2)
            self.geometry(f"{width}x{height}+{x}+{y}")
        except Exception:
            self.geometry(f"{width}x{height}")

        # Khung viền mạ vàng VKS
        border_frame = tk.Frame(self, bg="#C29B38", bd=2)
        border_frame.pack(fill="both", expand=True)

        inner_frame = tk.Frame(border_frame, bg="#FFFFFF")
        inner_frame.pack(fill="both", expand=True, padx=2, pady=2)

        # Thanh tiêu đề phụ bên trong
        banner = tk.Frame(inner_frame, bg="#07172C", padx=12, pady=7)
        banner.pack(fill="x")

        tk.Label(
            banner,
            text="🏛️  VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH",
            font=("Segoe UI", 9, "bold"),
            fg="#F59E0B",
            bg="#07172C",
        ).pack(side="left")

        # Nội dung chính
        content = tk.Frame(inner_frame, bg="#FFFFFF", padx=20, pady=12)
        content.pack(fill="both", expand=True)

        self.lbl_msg = tk.Label(
            content,
            text=message,
            font=("Segoe UI", 10, "bold"),
            fg="#0B2545",
            bg="#FFFFFF",
            wraplength=430,
            justify="center",
        )
        self.lbl_msg.pack(pady=(4, 8))

        # Thanh Progress bar chuyển động liên tục
        self.pb = ttk.Progressbar(content, mode="indeterminate", length=410)
        self.pb.pack(pady=4)
        self.pb.start(10)

        # Dòng trạng thái chi tiết kèm thời gian
        self.start_time = time.time()
        self.lbl_sub = tk.Label(
            content,
            text="⏳ Đang khởi tạo tiến trình... Vui lòng không đóng phần mềm.",
            font=("Segoe UI", 8, "italic"),
            fg="#64748B",
            bg="#FFFFFF",
        )
        self.lbl_sub.pack(pady=(6, 0))

        # Khóa tương tác với cửa sổ cha (Modal)
        self.protocol("WM_DELETE_WINDOW", lambda: None)
        try:
            self.grab_set()
        except Exception:
            pass
        self.focus_set()
        try:
            self.lift()
            self.attributes("-topmost", True)
        except Exception:
            pass

        self._ticker_active = True
        self._tick()

    def _tick(self):
        if not self._ticker_active:
            return
        elapsed = int(time.time() - self.start_time)
        mins, secs = divmod(elapsed, 60)
        cur_text = self.lbl_sub.cget("text")
        if "Đang chạy:" in cur_text:
            base_txt = cur_text.split(" (Đang chạy:")[0]
            self.lbl_sub.config(text=f"{base_txt} (Đang chạy: {mins:02d}:{secs:02d})")
        self.after(1000, self._tick)

    def update_message(self, text, subtext=None):
        try:
            self.lbl_msg.config(text=text)
            if subtext:
                elapsed = int(time.time() - self.start_time)
                mins, secs = divmod(elapsed, 60)
                self.lbl_sub.config(text=f"⚙️ {subtext} (Đang chạy: {mins:02d}:{secs:02d})")
            self.update_idletasks()
        except Exception:
            pass

    def close(self):
        self._ticker_active = False
        try:
            self.pb.stop()
            self.grab_release()
        except Exception:
            pass
        try:
            self.destroy()
        except Exception:
            pass


class ReconciliationDashboardDialog(tk.Toplevel):
    """
    Dashboard trực quan hiển thị kết quả tổng hợp & đối chiếu đa tập tin QLAHS.
    Thay thế hoàn toàn thông báo text messagebox đơn giản trước đây.
    """
    def __init__(self, parent, mode, records, match_n, file_stats, on_export_cb=None, on_view_table_cb=None, on_view_pivot_cb=None):
        super().__init__(parent)
        self.parent = parent
        self.mode = mode  # 'TIN_BAO' hoặc 'CASE'
        self.records = records or []
        self.match_n = match_n or 0
        self.file_stats = file_stats or []
        self.on_export_cb = on_export_cb
        self.on_view_table_cb = on_view_table_cb
        self.on_view_pivot_cb = on_view_pivot_cb

        self.title("⚡ DASHBOARD TỔNG HỢP & ĐỐI CHIẾU ĐA NGUỒN QLAHS")
        self.transient(parent)
        self.configure(bg="#0B2545")

        # Kích thước chuẩn Dashboard
        width, height = 980, 680
        self.minsize(860, 580)
        try:
            parent.update_idletasks()
            pw = parent.winfo_width()
            ph = parent.winfo_height()
            px = parent.winfo_rootx()
            py = parent.winfo_rooty()
            x = px + max(0, (pw - width) // 2)
            y = py + max(0, (ph - height) // 2)
            self.geometry(f"{width}x{height}+{x}+{y}")
        except Exception:
            self.geometry(f"{width}x{height}")

        icon_path = get_resource_path("logo_vks.ico")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        self._build_ui()
        self.grab_set()
        self.focus_set()

    def _build_ui(self):
        border_frame = tk.Frame(self, bg="#C29B38", bd=2)
        border_frame.pack(fill="both", expand=True)

        inner_frame = tk.Frame(border_frame, bg="#F8FAFC")
        inner_frame.pack(fill="both", expand=True, padx=2, pady=2)

        # 1. Header Banner
        banner = tk.Frame(inner_frame, bg="#07172C", padx=16, pady=10)
        banner.pack(fill="x")

        b_left = tk.Frame(banner, bg="#07172C")
        b_left.pack(side="left", fill="both", expand=True)

        tk.Label(
            b_left,
            text="🏛️  VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH",
            font=("Segoe UI", 9, "bold"),
            fg="#FBBF24",
            bg="#07172C"
        ).pack(anchor="w")

        is_multi = (len(self.file_stats) > 1)
        dash_title = (
            "⚡ DASHBOARD BÁO CÁO TỔNG KẾT RÀ SOÁT, ĐỐI CHIẾU & HỢP NHẤT ĐA NGUỒN QLAHS"
            if is_multi else
            "⚡ DASHBOARD BÁO CÁO KẾT QUẢ CHUẨN HÓA & ĐỐI CHIẾU DỮ LIỆU QLAHS"
        )
        tk.Label(
            b_left,
            text=dash_title,
            font=("Segoe UI", 12 if is_multi else 13, "bold"),
            fg="#FFFFFF",
            bg="#07172C"
        ).pack(anchor="w", pady=(2, 0))

        # Mode Badge bên phải
        if self.mode == "CASE_ENRICHED":
            badge_text = "🔗 LIÊN THÔNG VỤ ÁN ⟷ BỊ CAN (5 SHEET CHUYÊN NGHIỆP)"
            badge_bg = "#1E3A8A"
            badge_fg = "#FBBF24"
            unit_name = "Vụ án"
        elif self.mode == "TIN_BAO":
            badge_text = "🏷️ NGUỒN TIN VỀ TỘI PHẠM (36 CỘT CHUẨN)"
            badge_bg = "#0C4A6E"
            badge_fg = "#38BDF8"
            unit_name = "Tin báo"
        elif self.mode == "DEFENDANT":
            badge_text = "🏷️ DANH SÁCH BỊ CAN / TRUY NÃ / ĐÌNH NÃ (12 CỘT CHUẨN)"
            badge_bg = "#581C87"
            badge_fg = "#E9D5FF"
            unit_name = "Bị can"
        else:
            badge_text = "🏷️ HỒ SƠ VỤ ÁN HÌNH SỰ (27 CỘT CHUẨN)"
            badge_bg = "#064E3B"
            badge_fg = "#34D399"
            unit_name = "Vụ án"

        badge_frame = tk.Frame(banner, bg=badge_bg, padx=10, pady=6, bd=1, relief="solid")
        badge_frame.pack(side="right", padx=(10, 0))
        tk.Label(
            badge_frame,
            text=badge_text,
            font=("Segoe UI", 9, "bold"),
            fg=badge_fg,
            bg=badge_bg
        ).pack()

        # 2. Main Content
        content = tk.Frame(inner_frame, bg="#F8FAFC", padx=16, pady=12)
        content.pack(fill="both", expand=True)

        # Chỉ số tính toán
        total_merged = len(self.records)
        total_input_rows = sum(s.get("total", 0) for s in self.file_stats)
        match_count = self.match_n
        pure_new = max(0, total_merged - match_count)
        cross_rate = (match_count / total_input_rows * 100) if total_input_rows > 0 else 0.0

        # KPI CARDS FRAME
        cards_frame = tk.Frame(content, bg="#F8FAFC")
        cards_frame.pack(fill="x", pady=(0, 12))
        cards_frame.columnconfigure((0, 1, 2, 3), weight=1, uniform="cards")

        def make_card(parent, col, title, value, subtext, text_c, bg_c, border_c):
            cf = tk.Frame(parent, bg=bg_c, highlightthickness=1, highlightbackground=border_c, padx=12, pady=8)
            cf.grid(row=0, column=col, sticky="nsew", padx=4)
            tk.Label(cf, text=title, font=("Segoe UI", 8, "bold"), fg=text_c, bg=bg_c).pack(anchor="w")
            tk.Label(cf, text=value, font=("Segoe UI", 18, "bold"), fg=text_c, bg=bg_c).pack(anchor="w", pady=(2, 1))
            tk.Label(cf, text=subtext, font=("Segoe UI", 8), fg="#64748B", bg=bg_c).pack(anchor="w")

        if self.mode == "CASE_ENRICHED":
            linked_cnt = getattr(self.parent, "linked_defs_count", 0)
            unlinked_cnt = getattr(self.parent, "unlinked_defs_count", 0)
            make_card(cards_frame, 0, "TỔNG SỐ VỤ ÁN", f"{total_merged:,}", "Hồ sơ vụ án thụ lý", "#1E3A8A", "#EFF6FF", "#93C5FD")
            make_card(cards_frame, 1, "VỤ ÁN ĐÃ BỔ SUNG BỊ CAN", f"{match_count:,}", f"Khớp nối {linked_cnt:,} bị can", "#166534", "#F0FDF4", "#86EFAC")
            make_card(cards_frame, 2, "BỊ CAN ĐÃ LIÊN THÔNG", f"{linked_cnt:,}", "100% chính xác hồ sơ", "#0E7490", "#ECFEFF", "#A5F3FC")
            make_card(cards_frame, 3, "BỊ CAN CHỜ XÁC MINH (PM CŨ)", f"{unlinked_cnt:,}", "Mã cũ chờ đối soát", "#991B1B", "#FEF2F2", "#FECACA")
        else:
            card0_title = f"TỔNG {unit_name.upper()} HỢP NHẤT" if is_multi else f"TỔNG {unit_name.upper()} CHUẨN HÓA"
            make_card(cards_frame, 0, card0_title, f"{total_merged:,}", "100% dữ liệu thống nhất", "#1E3A8A", "#EFF6FF", "#93C5FD")

            sub_card1 = "Bổ sung thông tin còn thiếu" if match_count > 0 else "Chuẩn hóa & làm sạch 100%"
            make_card(cards_frame, 1, "KHỚP NỐI & BÙ ĐẮP DỮ LIỆU" if is_multi else "DỮ LIỆU ĐÃ LÀM SẠCH", f"{match_count:,}" if is_multi else f"{total_merged:,}", sub_card1, "#166534", "#F0FDF4", "#86EFAC")

            sub_card2 = f"Khớp chéo {cross_rate:.1f}%" if is_multi else "100% bản ghi hợp lệ"
            make_card(cards_frame, 2, "BẢN GHI ĐỘC LẬP / MỚI" if is_multi else "BẢN GHI HỢP LỆ", f"{pure_new:,}" if is_multi else f"{total_merged:,}", sub_card2, "#B45309", "#FFFBEB", "#FDE68A")

            sub_card3 = f"Tổng {total_input_rows:,} dòng nạp vào" if is_multi else "1 Tập tin dữ liệu gốc"
            make_card(cards_frame, 3, "TỆP TIN QLAHS ĐÃ RÀ SOÁT", f"{len(self.file_stats):,} Tệp", sub_card3, "#6B21A8", "#FAF5FF", "#D8B4FE")

        # SECTION BẢNG TẬP TIN NGUỒN
        tbl_lbl_frame = tk.Frame(content, bg="#F8FAFC")
        tbl_lbl_frame.pack(fill="x", pady=(0, 4))
        tk.Label(
            tbl_lbl_frame,
            text="📋 CHI TIẾT TỪNG TẬP TIN NGUỒN VÀ KẾT QUẢ ĐỐI CHIẾU CHÉO:",
            font=("Segoe UI", 10, "bold"),
            fg="#0F172A",
            bg="#F8FAFC"
        ).pack(side="left")

        # Treeview
        tree_frame = tk.Frame(content, bg="#FFFFFF", bd=1, relief="solid")
        tree_frame.pack(fill="both", expand=True, pady=(0, 10))

        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical")
        v_scroll.pack(side="right", fill="y")

        cols = ("stt", "fname", "total", "new", "matched", "contrib", "status")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", yscrollcommand=v_scroll.set, height=7)
        v_scroll.config(command=self.tree.yview)
        self.tree.pack(fill="both", expand=True)

        self.tree.heading("stt", text="STT")
        self.tree.heading("fname", text="Tên tập tin nguồn QLAHS")
        self.tree.heading("total", text="Tổng dòng nạp")
        self.tree.heading("new", text="Bản ghi mới")
        self.tree.heading("matched", text="Khớp nối bù đắp")
        self.tree.heading("contrib", text="Tỷ lệ đóng góp")
        self.tree.heading("status", text="Trạng thái xử lý")

        self.tree.column("stt", width=45, anchor="center")
        self.tree.column("fname", width=320, anchor="w")
        self.tree.column("total", width=100, anchor="center")
        self.tree.column("new", width=100, anchor="center")
        self.tree.column("matched", width=110, anchor="center")
        self.tree.column("contrib", width=105, anchor="center")
        self.tree.column("status", width=130, anchor="center")

        self.tree.tag_configure("odd", background="#F8FAFC")
        self.tree.tag_configure("even", background="#FFFFFF")
        self.tree.tag_configure("summary", background="#FEF3C7", font=("Segoe UI", 9, "bold"))

        for idx, s in enumerate(self.file_stats, 1):
            fn = s.get("fname", "")
            tot = s.get("total", 0)
            nw = s.get("new", 0)
            mt = s.get("matched", 0)
            pct = (nw / total_merged * 100) if total_merged > 0 else 0.0
            tag = "even" if idx % 2 == 0 else "odd"
            self.tree.insert("", "end", values=(
                idx, fn, f"{tot:,}", f"{nw:,}", f"{mt:,}", f"{pct:.1f}%", "✔ Đã đối chiếu & bù đắp"
            ), tags=(tag,))

        # Footer tổng kết
        self.tree.insert("", "end", values=(
            "Σ", f"TỔNG CỘNG CHUNG ({len(self.file_stats)} TẬP TIN)", f"{total_input_rows:,}", f"{total_merged:,}", f"{match_count:,}", "100.0%", "✔ Hoàn tất 100%"
        ), tags=("summary",))

        # NOTICE BOX
        notice_frame = tk.Frame(content, bg="#ECFDF5", highlightthickness=1, highlightbackground="#A7F3D0", padx=12, pady=8)
        notice_frame.pack(fill="x", pady=(0, 4))

        tk.Label(
            notice_frame,
            text="★ KẾT QUẢ CHUẨN HÓA DỮ LIỆU ĐÃ THỰC HIỆN TỰ ĐỘNG:",
            font=("Segoe UI", 9, "bold"),
            fg="#047857",
            bg="#ECFDF5"
        ).pack(anchor="w")

        if self.mode == "CASE_ENRICHED":
            notes = [
                "• Đã tự động đối chiếu và liên thông dữ liệu giữa Sổ quản lý vụ án và Danh sách bị can QLAHS.",
                "• Bổ sung chính xác danh sách chi tiết bị can, số lượng bị can và bị can đầu vụ vào từng vụ án tương ứng.",
                "• Bị can mang mã cũ hoặc vụ án ngoài danh sách được gán 'Xác minh sau' và xuất sang Sheet riêng để đối soát tiếp từ PM cũ."
            ]
        else:
            notes = [
                "• Đã phân tách độc lập các cột nghiệp vụ: Lãnh đạo Viện, Lãnh đạo Phòng, Kiểm sát viên, Cán bộ thụ lý; xóa sạch tiền tố PVT, Đ/c...",
                "• Toàn bộ ngày tháng đã chuẩn hóa định dạng DD/MM/YYYY; tự động tính toán số ngày còn lại và cảnh báo quá hạn thời hạn điều tra.",
                "• Đã bù đắp chéo dữ liệu còn thiếu giữa các tệp tin để tạo ra bộ dữ liệu sạch, hoàn chỉnh và có giá trị sử dụng cao nhất."
            ]
        for note in notes:
            tk.Label(notice_frame, text=note, font=("Segoe UI", 8), fg="#065F46", bg="#ECFDF5").pack(anchor="w", padx=(10, 0))

        # 3. ACTION BUTTON BAR (BOTTOM)
        act_bar = tk.Frame(inner_frame, bg="#07172C", padx=16, pady=10)
        act_bar.pack(fill="x", side="bottom")

        btn_export_text = "📥 Xuất File Excel Liên Thông 5 Sheet (.xlsx)" if self.mode == "CASE_ENRICHED" else "📥 Xuất File Excel Đầy Đủ & Pivot (.xlsx)"
        btn_export = tk.Button(
            act_bar,
            text=btn_export_text,
            font=("Segoe UI", 10, "bold"),
            bg="#15803D",
            fg="#FFFFFF",
            activebackground="#166534",
            activeforeground="#FFFFFF",
            relief="flat",
            cursor="hand2",
            padx=14,
            pady=6,
            command=self._on_export_click
        )
        btn_export.pack(side="left", padx=(0, 8))

        btn_table = tk.Button(
            act_bar,
            text="📋 Xem Trên Bảng Dữ Liệu",
            font=("Segoe UI", 10, "bold"),
            bg="#1D4ED8",
            fg="#FFFFFF",
            activebackground="#1E40AF",
            activeforeground="#FFFFFF",
            relief="flat",
            cursor="hand2",
            padx=14,
            pady=6,
            command=self._on_view_table_click
        )
        btn_table.pack(side="left", padx=(0, 8))

        btn_pivot = tk.Button(
            act_bar,
            text="📈 Xem Bảng Ma Trận Pivot Table",
            font=("Segoe UI", 10, "bold"),
            bg="#D97706",
            fg="#FFFFFF",
            activebackground="#B45309",
            activeforeground="#FFFFFF",
            relief="flat",
            cursor="hand2",
            padx=14,
            pady=6,
            command=self._on_view_pivot_click
        )
        btn_pivot.pack(side="left", padx=(0, 8))

        btn_close = tk.Button(
            act_bar,
            text="❌ Đóng Dashboard",
            font=("Segoe UI", 10),
            bg="#475569",
            fg="#FFFFFF",
            activebackground="#334155",
            activeforeground="#FFFFFF",
            relief="flat",
            cursor="hand2",
            padx=14,
            pady=6,
            command=self.destroy
        )
        btn_close.pack(side="right")

        tk.Label(
            self,
            text="⚡ Phát triển bởi: Trần Văn Khoa - Kiểm tra viên / VKSND TP.HCM",
            font=("Segoe UI", 8, "italic"),
            fg="#94A3B8",
            bg="#0B2545"
        ).pack(side="bottom", pady=(2, 6))

    def _on_export_click(self):
        if self.on_export_cb:
            self.on_export_cb()

    def _on_view_table_click(self):
        if self.on_view_table_cb:
            self.on_view_table_cb()
        self.destroy()

    def _on_view_pivot_click(self):
        if self.on_view_pivot_cb:
            self.on_view_pivot_cb()
        self.destroy()


# ==============================================================================
# HÀM VẼ BIỂU ĐỒ TIẾN ĐỘ TRỰC QUAN (VISUAL PROGRESS & DISTRIBUTION CANVAS)
# ==============================================================================
def draw_visual_distribution_chart(canvas, width, height, segments):
    """
    Vẽ thanh phân bổ tỷ lệ trực quan (Stacked Horizontal Distribution Bar)
    segments: danh sách các tuple (nhãn, số_lượng, mã_màu)
    """
    try:
        canvas.delete("all")
        total = sum(c for _, c, _ in segments)
        if total <= 0:
            canvas.create_text(width // 2, height // 2, text="Chưa có dữ liệu thống kê tiến độ", fill="#64748B", font=("Segoe UI", 9, "italic"))
            return

        pad_x = 8
        bar_y1 = 6
        bar_y2 = 22
        usable_w = max(10, width - (2 * pad_x))

        cur_x = pad_x
        for lbl, count, color in segments:
            if count <= 0:
                continue
            seg_w = max(3, int((count / total) * usable_w))
            canvas.create_rectangle(cur_x, bar_y1, cur_x + seg_w, bar_y2, fill=color, outline="#FFFFFF", width=1)
            cur_x += seg_w

        # Vẽ Chú thích (Legend)
        leg_x = pad_x
        leg_y = bar_y2 + 8
        for lbl, count, color in segments:
            pct = (count / total * 100) if total else 0.0
            txt = f"{lbl}: {count:,} ({pct:.1f}%)"
            canvas.create_oval(leg_x, leg_y + 2, leg_x + 9, leg_y + 11, fill=color, outline="")
            t_id = canvas.create_text(leg_x + 13, leg_y + 6, text=txt, anchor="w", font=("Segoe UI", 8), fill="#1E293B")
            bbox = canvas.bbox(t_id)
            leg_x = (bbox[2] + 16) if bbox else (leg_x + 110)
    except Exception:
        pass


# ==============================================================================
# BÁO CÁO TRÍCH YẾU NGHIỆP VỤ (EXECUTIVE BRIEF GENERATOR)
# ==============================================================================
def generate_executive_brief_text(records, data_format="CASE", config=None):
    """
    Tự động sinh đoạn văn bản trích yếu số liệu báo cáo nghiệp vụ chuẩn văn phong ngành Kiểm sát.
    Sẵn sàng sao chép nhanh (Ctrl+C) hoặc lưu tệp Word/Text báo cáo Lãnh đạo Viện / Phòng.
    """
    cfg = config or (CURRENT_CONFIG if "CURRENT_CONFIG" in globals() else DEFAULT_APP_CONFIG)
    agency = cfg.get("agency_name", "VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH")
    author = cfg.get("author_title", "Trần Văn Khoa - Kiểm tra viên / VKSND TP.HCM")
    thresholds = cfg.get("warning_thresholds", {"urgent_days": 30, "very_urgent_days": 60, "warning_days": 90})

    now_str = datetime.now().strftime("%d/%m/%Y lúc %H:%M")
    is_tin_bao = (data_format == "TIN_BAO")
    label_type = "Nguồn tin về tội phạm (Tin báo & Tố giác)" if is_tin_bao else "Hồ sơ vụ án hình sự"

    total = len(records)
    if total == 0:
        return f"{agency}\nBÁO CÁO TRÍCH YẾU NGHIỆP VỤ\n(Chưa có dữ liệu nào được nạp vào phần mềm)"

    # Phân loại tiến độ & trạng thái
    resolved = sum(1 for r in records if any(k in str(r.get("ket_qua_giai_quyet") or r.get("trang_thai_vu_an") or r.get("trang_thai") or "").lower() for k in ["đã giải quyết", "kết luận", "đình chỉ", "khởi tố"]))
    suspended = sum(1 for r in records if any(k in str(r.get("ket_qua_giai_quyet") or r.get("muc_do_khan_cap") or r.get("trang_thai_vu_an") or "").lower() for k in ["tạm đình chỉ", "đình nã", "truy nã"]))
    in_progress = max(0, total - resolved - suspended)

    # Thống kê cảnh báo thời hiệu & thời hạn
    overdue = sum(1 for r in records if any("quá hạn" in str(r.get(f) or "").lower() for f in ["muc_do_khan_cap", "canh_bao_thoi_hieu"]))
    red = sum(1 for r in records if any("🔴" in str(r.get(f) or "") for f in ["muc_do_khan_cap", "canh_bao_thoi_hieu"]))
    orange = sum(1 for r in records if any("🟠" in str(r.get(f) or "") for f in ["muc_do_khan_cap", "canh_bao_thoi_hieu"]))
    yellow = sum(1 for r in records if any("🟡" in str(r.get(f) or "") for f in ["muc_do_khan_cap", "canh_bao_thoi_hieu"]))
    green = sum(1 for r in records if any("🟢" in str(r.get(f) or "") for f in ["muc_do_khan_cap", "canh_bao_thoi_hieu"]))

    # Danh sách các hồ sơ khẩn cấp (Overdue, Red, Orange)
    urgent_records = [
        r for r in records
        if any(sym in (str(r.get("muc_do_khan_cap") or "") + " " + str(r.get("canh_bao_thoi_hieu") or "")).lower() for sym in ["🔴", "quá hạn", "khẩn cấp"])
    ]

    # Phân bổ theo KSV
    ksv_map = {}
    for r in records:
        ksv = r.get("ksv_chinh") or r.get("can_bo_giai_quyet") or "Chưa phân công KSV"
        ksv_map[ksv] = ksv_map.get(ksv, 0) + 1
    top_ksv = sorted(ksv_map.items(), key=lambda x: x[1], reverse=True)[:8]

    # Phân bổ theo Đơn vị VKS thụ lý
    unit_map = {}
    for r in records:
        u = r.get("vks_thu_ly") or "Chưa rõ đơn vị"
        unit_map[u] = unit_map.get(u, 0) + 1
    top_units = sorted(unit_map.items(), key=lambda x: x[1], reverse=True)

    lines = []
    lines.append(agency.upper())
    lines.append("=" * 78)
    lines.append("BÁO CÁO TRÍCH YẾU TỔNG HỢP TIẾN ĐỘ & THỜI HIỆU TRUY CỨU TRÁCH NHIỆM HÌNH SỰ")
    lines.append(f"(Số liệu tổng hợp tự động từ phần mềm QLAHS tính đến {now_str})")
    lines.append("-" * 78)
    lines.append("")

    lines.append(f"I. TỔNG QUAN TÌNH HÌNH QUẢN LÝ {label_type.upper()}:")
    lines.append(f"  • Tổng số hồ sơ đang quản lý theo dõi: {total:,} hồ sơ (100.0%)")
    lines.append(f"  • Hồ sơ đang trong thời hạn giải quyết: {in_progress:,} hồ sơ ({in_progress / total * 100:.1f}%)")
    lines.append(f"  • Hồ sơ đã giải quyết / kết luận xử lý: {resolved:,} hồ sơ ({resolved / total * 100:.1f}%)")
    lines.append(f"  • Hồ sơ đang tạm đình chỉ giải quyết: {suspended:,} hồ sơ ({suspended / total * 100:.1f}%)")
    if not is_tin_bao:
        stg_map = {}
        for r in records:
            gd = r.get("giai_doan", "").strip() or "Chưa rõ"
            stg_map[gd] = stg_map.get(gd, 0) + 1
        stg_parts = []
        for s in ("Điều tra", "Truy tố", "Sơ thẩm", "Phúc thẩm", "Thi hành án"):
            if s in stg_map:
                stg_parts.append(f"{s}: {stg_map[s]:,} vụ ({stg_map[s] / total * 100:.1f}%)")
        for s, cnt in stg_map.items():
            if s not in ("Điều tra", "Truy tố", "Sơ thẩm", "Phúc thẩm", "Thi hành án", "Chưa rõ"):
                stg_parts.append(f"{s}: {cnt:,} vụ ({cnt / total * 100:.1f}%)")
        if stg_parts:
            lines.append(f"  • Phân bổ Giai đoạn tố tụng: " + " | ".join(stg_parts))
    lines.append("")

    lines.append("II. THEO DÕI TIẾN ĐỘ & THỜI HIỆU TRUY CỨU TNHS (ĐẶC BIỆT CHÚ Ý):")
    lines.append(f"  • ⚠️ Đã quá thời hạn giải quyết / Quá thời hiệu TNHS: {overdue:,} hồ sơ")
    lines.append(f"  • 🔴 Khẩn cấp / Cấp bách (Thời hạn còn ≤ {thresholds.get('urgent_days', 30)} ngày): {red:,} hồ sơ")
    lines.append(f"  • 🟠 Rất gấp (Thời hạn còn 31 - {thresholds.get('very_urgent_days', 60)} ngày): {orange:,} hồ sơ")
    lines.append(f"  • 🟡 Sắp hết hạn (Thời hạn còn 61 - {thresholds.get('warning_days', 90)} ngày): {yellow:,} hồ sơ")
    lines.append(f"  • 🟢 Trong hạn bình thường (Thời hạn còn > {thresholds.get('warning_days', 90)} ngày): {green:,} hồ sơ")
    lines.append("")

    if urgent_records:
        lines.append(f"III. DANH SÁCH {min(len(urgent_records), 20)} HỒ SƠ KHẨN CẤP / QUÁ HẠN CẦN TẬP TRUNG CHỈ ĐẠO XỬ LÝ:")
        for idx, r in enumerate(urgent_records[:20], 1):
            code = r.get("ma_tin_bao") or r.get("ma_vu_an") or f"STT {idx}"
            title = r.get("noi_dung_tin_bao") or r.get("ten_vu_an") or "Chưa có tên/nội dung"
            title = (title[:75] + "...") if len(title) > 75 else title
            ksv = r.get("ksv_chinh") or r.get("can_bo_giai_quyet") or "Chưa phân công"
            unit = r.get("vks_thu_ly") or ""
            due = r.get("thoi_han_den_ngay") or r.get("ngay_het_thoi_hieu") or "Chưa rõ"
            urg = r.get("muc_do_khan_cap") or r.get("canh_bao_thoi_hieu") or "Khẩn cấp"
            lines.append(f"  {idx}. [{code}] {title}")
            lines.append(f"     -> Đơn vị: {unit} | KSV: {ksv}")
            lines.append(f"     -> Hạn / Thời hiệu: {due} | Tiến độ: {urg}")
        if len(urgent_records) > 20:
            lines.append(f"  ... và còn {len(urgent_records) - 20} hồ sơ khẩn cấp khác (xem chi tiết tại bảng xuất Excel).")
        lines.append("")

    if top_units:
        lines.append("IV. CƠ CẤU HỒ SƠ THEO ĐƠN VI VKS THỤ LÝ:")
        for u_name, u_cnt in top_units[:10]:
            lines.append(f"  • {u_name}: {u_cnt:,} hồ sơ ({u_cnt / total * 100:.1f}%)")
        lines.append("")

    if top_ksv:
        lines.append("V. TÌNH HÌNH PHÂN CÔNG THEO KIỂM SÁT VIÊN THỤ LÝ (TOP 8):")
        for ksv_name, ksv_cnt in top_ksv:
            lines.append(f"  • KSV {ksv_name}: {ksv_cnt:,} hồ sơ ({ksv_cnt / total * 100:.1f}%)")
        lines.append("")

    lines.append("=" * 78)
    lines.append(f"Người lập trích yếu báo cáo: {author}")
    lines.append("Phần Mềm Chuẩn Hóa & Đối Chiếu Dữ Liệu QLAHS (Version 1.0 Offline)")
    return "\n".join(lines)


# ==============================================================================
# HỆ THỐNG SO SÁNH BIẾN ĐỘNG 2 KỲ BÁO CÁO (DIFF ANALYSIS ENGINE)
# ==============================================================================
def compare_two_qlahs_periods(records_old, records_new, data_format=None):
    """
    So sánh biến động dữ liệu giữa 2 kỳ báo cáo QLAHS:
    - added: Danh sách hồ sơ mới phát sinh trong kỳ
    - resolved: Danh sách hồ sơ đã được giải quyết / kết thúc
    - worsened: Danh sách hồ sơ chuyển mức khẩn cấp về thời hạn / thời hiệu TNHS
    - continued: Danh sách hồ sơ tiếp tục tồn đọng / đang giải quyết
    """
    def norm_k(s):
        return re.sub(r'[^a-zA-Z0-9]', '', str(s or '')).upper()

    def get_urgency_level(r):
        u1 = str(r.get("muc_do_khan_cap") or "").lower()
        u2 = str(r.get("canh_bao_thoi_hieu") or "").lower()
        cb = u1 + " " + u2
        if "quá hạn" in cb:
            return 5
        if "🔴" in cb:
            return 4
        if "🟠" in cb:
            return 3
        if "🟡" in cb:
            return 2
        if "🟢" in cb:
            return 1
        return 0

    old_by_code = {}
    old_by_sub = {}
    for r in records_old:
        code = norm_k(r.get("ma_tin_bao") or r.get("ma_vu_an"))
        if code:
            old_by_code[code] = r
        sub = norm_k(str(r.get("so_qd_khoi_to") or r.get("so_thu_ly") or "") + str(r.get("vks_thu_ly") or ""))
        if sub:
            old_by_sub[sub] = r

    matched_old_keys = set()
    added = []
    worsened = []
    continued = []

    for r_new in records_new:
        c_k = norm_k(r_new.get("ma_tin_bao") or r_new.get("ma_vu_an"))
        s_k = norm_k(str(r_new.get("so_qd_khoi_to") or r_new.get("so_thu_ly") or "") + str(r_new.get("vks_thu_ly") or ""))

        r_old = None
        if c_k and c_k in old_by_code:
            r_old = old_by_code[c_k]
            matched_old_keys.add(("code", c_k))
        elif s_k and s_k in old_by_sub:
            r_old = old_by_sub[s_k]
            matched_old_keys.add(("sub", s_k))

        if r_old is None:
            added.append(r_new)
        else:
            lvl_old = get_urgency_level(r_old)
            lvl_new = get_urgency_level(r_new)
            if lvl_new > lvl_old and lvl_new >= 2:
                rec_w = dict(r_new)
                rec_w["_old_urgency"] = r_old.get("muc_do_khan_cap") or r_old.get("canh_bao_thoi_hieu") or "Bình thường"
                rec_w["_new_urgency"] = r_new.get("muc_do_khan_cap") or r_new.get("canh_bao_thoi_hieu") or "Khẩn cấp"
                worsened.append(rec_w)
            else:
                continued.append(r_new)

    resolved = []
    for r_old in records_old:
        c_k = norm_k(r_old.get("ma_tin_bao") or r_old.get("ma_vu_an"))
        s_k = norm_k(str(r_old.get("so_qd_khoi_to") or r_old.get("so_thu_ly") or "") + str(r_old.get("vks_thu_ly") or ""))
        is_m = False
        if c_k and ("code", c_k) in matched_old_keys:
            is_m = True
        elif s_k and ("sub", s_k) in matched_old_keys:
            is_m = True
        if not is_m:
            resolved.append(r_old)

    stats = {
        "total_old": len(records_old),
        "total_new": len(records_new),
        "n_added": len(added),
        "n_resolved": len(resolved),
        "n_worsened": len(worsened),
        "n_continued": len(continued),
        "net_change": len(records_new) - len(records_old)
    }

    return {
        "stats": stats,
        "added": added,
        "resolved": resolved,
        "worsened": worsened,
        "continued": continued,
        "data_format": data_format or "CASE"
    }


def export_diff_to_excel(diff_data, out_filepath, title_old="Kỳ Trước", title_new="Kỳ Này"):
    """
    Xuất báo cáo phân tích biến động giữa 2 kỳ thành tệp Excel chuyên nghiệp đa Sheet có Biểu đồ BarChart.
    """
    if openpyxl is None:
        raise RuntimeError("Thư viện openpyxl chưa được cài đặt.")

    wb = openpyxl.Workbook()
    stats = diff_data["stats"]
    is_tin_bao = (diff_data.get("data_format") == "TIN_BAO")
    cols_def = COLUMNS_TIN_BAO_35 if is_tin_bao else COLUMNS_CASE_24

    font_title = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
    font_bold = Font(name="Segoe UI", size=9, bold=True, color="1E293B")
    font_regular = Font(name="Segoe UI", size=9)
    thin_border = Border(left=Side(style="thin", color="CBD5E1"), right=Side(style="thin", color="CBD5E1"),
                         top=Side(style="thin", color="CBD5E1"), bottom=Side(style="thin", color="CBD5E1"))

    # Sheet 1: Tổng quan biến động
    ws1 = wb.active
    ws1.title = "Tong_Quan_Bien_Dong"
    ws1.views.sheetView[0].showGridLines = True

    ws1.merge_cells("A1:G1")
    t1 = ws1["A1"]
    t1.value = "BÁO CÁO PHÂN TÍCH SO SÁNH BIẾN ĐỘNG GIỮA 2 KỲ BÁO CÁO QLAHS"
    t1.font = font_title
    t1.fill = PatternFill(start_color="0A2540", end_color="0A2540", fill_type="solid")
    t1.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 32

    ws1.merge_cells("A2:G2")
    t2 = ws1["A2"]
    t2.value = f"Kỳ trước: {title_old}  |  Kỳ này: {title_new}  •  Thời điểm đối chiếu: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    t2.font = Font(name="Segoe UI", size=9, italic=True, color="475569")
    t2.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[2].height = 20

    # Bảng chỉ số
    headers_stat = ["STT", "Chỉ Tiêu So Sánh Biến Động Nghiệp Vụ", "Số Lượng", "Tỷ Lệ %", "Đánh Giá Chiều Hướng Biến Động"]
    ws1.row_dimensions[4].height = 24
    for i, h in enumerate(headers_stat, 1):
        c = ws1.cell(row=4, column=i, value=h)
        c.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")

    tot_n = stats["total_new"] or 1
    stat_rows = [
        ("1", f"Tổng số hồ sơ Kỳ trước ({title_old})", stats["total_old"], "-", "Mốc đối chiếu ban đầu"),
        ("2", f"Tổng số hồ sơ Kỳ này ({title_new})", stats["total_new"], "100.0%", "Hiện trạng theo dõi hiện tại"),
        ("3", "Chênh lệch ròng (Tăng / Giảm)", f"{'+' if stats['net_change'] >= 0 else ''}{stats['net_change']}", f"{abs(stats['net_change']) / (stats['total_old'] or 1) * 100:.1f}%", "Xu hướng khối lượng công việc"),
        ("4", "✨ Số hồ sơ mới thụ lý / Mới phát sinh trong kỳ", stats["n_added"], f"{stats['n_added'] / tot_n * 100:.1f}%", "🟢 Thụ lý mới"),
        ("5", "✅ Số hồ sơ đã giải quyết xong / Kết thúc trong kỳ", stats["n_resolved"], f"{stats['n_resolved'] / (stats['total_old'] or 1) * 100:.1f}%", "🔵 Đã giải quyết / Đạt kết quả"),
        ("6", "⚠️ Số hồ sơ chuyển mức khẩn cấp thời hiệu / Quá hạn", stats["n_worsened"], f"{stats['n_worsened'] / tot_n * 100:.1f}%", "🔴 Nguy cơ quá hạn thời hiệu"),
        ("7", "⏸ Số hồ sơ tiếp tục tồn đọng / Tạm đình chỉ kéo dài", stats["n_continued"], f"{stats['n_continued'] / tot_n * 100:.1f}%", "⚪ Cần tiếp tục theo dõi"),
    ]

    for r_i, (stt, name, val, pct, note) in enumerate(stat_rows, 5):
        ws1.row_dimensions[r_i].height = 20
        c1 = ws1.cell(row=r_i, column=1, value=stt)
        c2 = ws1.cell(row=r_i, column=2, value=name)
        c3 = ws1.cell(row=r_i, column=3, value=val)
        c4 = ws1.cell(row=r_i, column=4, value=pct)
        c5 = ws1.cell(row=r_i, column=5, value=note)
        for c in (c1, c2, c3, c4, c5):
            c.border = thin_border
            c.font = font_regular
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c2.alignment = Alignment(horizontal="left", vertical="center")
        c3.alignment = Alignment(horizontal="center", vertical="center")
        c4.alignment = Alignment(horizontal="center", vertical="center")
        c5.alignment = Alignment(horizontal="left", vertical="center")
        if r_i == 8:
            c2.fill = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
        elif r_i == 9:
            c2.fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
        elif r_i == 10 and stats["n_worsened"] > 0:
            c2.fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
            c3.font = Font(name="Segoe UI", size=9, bold=True, color="991B1B")

    ws1.column_dimensions["A"].width = 8
    ws1.column_dimensions["B"].width = 46
    ws1.column_dimensions["C"].width = 16
    ws1.column_dimensions["D"].width = 14
    ws1.column_dimensions["E"].width = 30

    # Nhúng biểu đồ cột BarChart trực quan
    if BarChart and Reference:
        try:
            chart = BarChart()
            chart.type = "col"
            chart.style = 10
            chart.title = "CƠ CẤU BIẾN ĐỘNG HỒ SƠ GIỮA 2 KỲ BÁO CÁO"
            chart.y_axis.title = "Số lượng hồ sơ"
            chart.x_axis.title = "Nhóm biến động"
            data_ref = Reference(ws1, min_col=3, min_row=7, max_row=11)
            cats_ref = Reference(ws1, min_col=2, min_row=8, max_row=11)
            chart.add_data(data_ref, titles_from_data=True)
            chart.set_categories(cats_ref)
            chart.width = 17
            chart.height = 10
            ws1.add_chart(chart, "G4")
        except Exception:
            pass

    # Hàm phụ ghi danh sách chi tiết từng sheet
    def write_diff_sheet(title, rec_list, header_color):
        ws = wb.create_sheet(title=title)
        ws.views.sheetView[0].showGridLines = True
        ws.row_dimensions[1].height = 26
        # Headers
        ws.cell(row=1, column=1, value="STT").alignment = Alignment(horizontal="center", vertical="center")
        for c_idx, (_, col_name, col_w) in enumerate(cols_def, start=1):
            cell = ws.cell(row=1, column=c_idx, value=col_name)
            cell.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.column_dimensions[get_column_letter(c_idx)].width = max(10, min(col_w // 8, 35))

        for r_idx, r in enumerate(rec_list, start=2):
            ws.row_dimensions[r_idx].height = 20
            for c_idx, (k, _, _) in enumerate(cols_def, start=1):
                val = r.get(k, "")
                if c_idx == 1:
                    val = r_idx - 1
                c = ws.cell(row=r_idx, column=c_idx, value=val)
                c.border = thin_border
                c.font = font_regular
                c.alignment = Alignment(horizontal="center" if c_idx in (1, 2) else "left", vertical="center")

    write_diff_sheet("1.Moi_Thu_Ly_Trong_Ky", diff_data["added"], "047857")
    write_diff_sheet("2.Da_Giai_Quyet_Trong_Ky", diff_data["resolved"], "1D4ED8")
    write_diff_sheet("3.Chuyen_Cap_Khan_Cap", diff_data["worsened"], "B91C1C")
    write_diff_sheet("4.Tiep_Tuc_Ton_Dong", diff_data["continued"][:500], "475569")

    wb.active = 0
    wb.save(out_filepath)
    return out_filepath


# ==============================================================================
# HỆ THỐNG CÁC CỬA SỔ HỘI THOẠI BỔ SUNG (DIALOGS)
# ==============================================================================
class SettingsDialog(tk.Toplevel):
    """
    Cửa sổ Cài đặt cấu hình tùy biến:
    - Tên cơ quan chủ quản
    - Chức danh tác giả
    - Ngưỡng cảnh báo thời hiệu (30, 60, 90 ngày)
    - Danh mục từ điển viết tắt đơn vị (Custom Aliases)
    """
    def __init__(self, parent, config=None, on_save_cb=None):
        super().__init__(parent)
        self.parent = parent
        self.config = dict(config or CURRENT_CONFIG)
        self.on_save_cb = on_save_cb

        self.title("⚙️ CÀI ĐẶT QUY TẮC CHUẨN HÓA & NGƯỠNG CẢNH BÁO QLAHS")
        self.geometry("780x560")
        self.minsize(700, 500)
        self.transient(parent)
        self.configure(bg="#F8FAFC")

        icon_path = get_resource_path("logo_vks.ico")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        self._build_ui()
        self.grab_set()
        self.focus_set()

    def _build_ui(self):
        # Header
        header = tk.Frame(self, bg="#07172C", padx=16, pady=10)
        header.pack(fill=tk.X)
        tk.Label(header, text="⚙️ CÀI ĐẶT HỆ THỐNG & TÙY BIẾN QUY TẮC NGHIỆP VỤ", font=("Segoe UI", 12, "bold"), fg="#FBBF24", bg="#07172C").pack(anchor="w")
        tk.Label(header, text="Tùy chỉnh tên cơ quan, ngưỡng cảnh báo thời hiệu và danh mục viết tắt đơn vị linh hoạt theo từng địa phương.", font=("Segoe UI", 8, "italic"), fg="#93C5FD", bg="#07172C").pack(anchor="w")

        # Notebook Tabs
        nb = ttk.Notebook(self)
        nb.pack(fill=tk.BOTH, expand=True, padx=12, pady=10)

        # Tab 1: Cơ quan & Tác giả
        tab1 = tk.Frame(nb, bg="#FFFFFF", padx=16, pady=16)
        nb.add(tab1, text="🏢 Cơ Quan & Tác Giả")

        tk.Label(tab1, text="Tên Cơ quan chủ quản (Xuất hiện trên tiêu đề & báo cáo):", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", fg="#1E293B").pack(anchor="w", pady=(0, 4))
        self.ent_agency = ttk.Entry(tab1, font=("Segoe UI", 10), width=65)
        self.ent_agency.insert(0, self.config.get("agency_name", DEFAULT_APP_CONFIG["agency_name"]))
        self.ent_agency.pack(anchor="w", pady=(0, 16))

        tk.Label(tab1, text="Thông tin tác giả / Chức danh người lập báo cáo:", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", fg="#1E293B").pack(anchor="w", pady=(0, 4))
        self.ent_author = ttk.Entry(tab1, font=("Segoe UI", 10), width=65)
        self.ent_author.insert(0, self.config.get("author_title", DEFAULT_APP_CONFIG["author_title"]))
        self.ent_author.pack(anchor="w", pady=(0, 16))

        info_box = tk.Frame(tab1, bg="#EFF6FF", padx=12, pady=10, relief=tk.SOLID, bd=1)
        info_box.pack(fill=tk.X, pady=(10, 0))
        tk.Label(info_box, text="💡 Ghi chú:", font=("Segoe UI", 9, "bold"), fg="#1D4ED8", bg="#EFF6FF").pack(anchor="w")
        tk.Label(info_box, text="Các thông tin này sẽ được tự động điền vào tiêu đề phần mềm, các bảng báo cáo Word/Text và tiêu đề các file Excel xuất ra.", font=("Segoe UI", 8), fg="#1E40AF", bg="#EFF6FF").pack(anchor="w")

        # Tab 2: Ngưỡng cảnh báo thời hiệu
        tab2 = tk.Frame(nb, bg="#FFFFFF", padx=16, pady=16)
        nb.add(tab2, text="⚠️ Ngưỡng Cảnh Báo Thời Hiệu")

        th = self.config.get("warning_thresholds", DEFAULT_APP_CONFIG["warning_thresholds"])
        tk.Label(tab2, text="THIẾT LẬP MỐC ĐẾM NGƯỢC THỜI HẠN & THỜI HIỆU TRUY CỨU TNHS:", font=("Segoe UI", 10, "bold"), bg="#FFFFFF", fg="#1E293B").pack(anchor="w", pady=(0, 12))

        # Urgent (Red)
        f_u = tk.Frame(tab2, bg="#FFFFFF")
        f_u.pack(anchor="w", pady=6)
        tk.Label(f_u, text="🔴 Mức KHẨN CẤP / CẤP BÁCH (Đỏ): Thời hạn còn ≤", font=("Segoe UI", 9, "bold"), fg="#991B1B", bg="#FFFFFF").pack(side=tk.LEFT)
        self.sp_urgent = ttk.Spinbox(f_u, from_=1, to=365, width=6, font=("Segoe UI", 9, "bold"))
        self.sp_urgent.set(th.get("urgent_days", 30))
        self.sp_urgent.pack(side=tk.LEFT, padx=6)
        tk.Label(f_u, text="ngày", font=("Segoe UI", 9), bg="#FFFFFF").pack(side=tk.LEFT)

        # Very Urgent (Orange)
        f_vu = tk.Frame(tab2, bg="#FFFFFF")
        f_vu.pack(anchor="w", pady=6)
        tk.Label(f_vu, text="🟠 Mức RẤT GẤP (Cam): Thời hạn còn ≤", font=("Segoe UI", 9, "bold"), fg="#C2410C", bg="#FFFFFF").pack(side=tk.LEFT)
        self.sp_vurgent = ttk.Spinbox(f_vu, from_=1, to=365, width=6, font=("Segoe UI", 9, "bold"))
        self.sp_vurgent.set(th.get("very_urgent_days", 60))
        self.sp_vurgent.pack(side=tk.LEFT, padx=6)
        tk.Label(f_vu, text="ngày", font=("Segoe UI", 9), bg="#FFFFFF").pack(side=tk.LEFT)

        # Warning (Yellow)
        f_w = tk.Frame(tab2, bg="#FFFFFF")
        f_w.pack(anchor="w", pady=6)
        tk.Label(f_w, text="🟡 Mức SẮP ĐẾN HẠN (Vàng): Thời hạn còn ≤", font=("Segoe UI", 9, "bold"), fg="#B45309", bg="#FFFFFF").pack(side=tk.LEFT)
        self.sp_warning = ttk.Spinbox(f_w, from_=1, to=365, width=6, font=("Segoe UI", 9, "bold"))
        self.sp_warning.set(th.get("warning_days", 90))
        self.sp_warning.pack(side=tk.LEFT, padx=6)
        tk.Label(f_w, text="ngày (Trên mốc này là 🟢 Bình thường)", font=("Segoe UI", 9), bg="#FFFFFF").pack(side=tk.LEFT)

        # Tab 3: Danh mục viết tắt đơn vị
        tab3 = tk.Frame(nb, bg="#FFFFFF", padx=12, pady=12)
        nb.add(tab3, text="🏷️ Viết Tắt Đơn Vị (Aliases)")

        tk.Label(tab3, text="DANH MỤC QUY TẮC NHẬN DIỆN VIẾT TẮT ĐƠN VỊ:", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", fg="#1E293B").pack(anchor="w", pady=(0, 6))

        tree_frame = tk.Frame(tab3, bg="#FFFFFF")
        tree_frame.pack(fill=tk.BOTH, expand=True)

        self.tree_alias = ttk.Treeview(tree_frame, columns=("alias", "target"), show="headings", height=8)
        self.tree_alias.heading("alias", text="Từ khóa viết tắt (Ví dụ: P2, KV4)")
        self.tree_alias.heading("target", text="Tên đơn vị chuẩn hóa thay thế (Ví dụ: Phòng 2 - VKSND TP.HCM)")
        self.tree_alias.column("alias", width=220, anchor=tk.CENTER)
        self.tree_alias.column("target", width=460, anchor=tk.W)

        sb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree_alias.yview)
        self.tree_alias.configure(yscrollcommand=sb.set)
        self.tree_alias.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        curr_aliases = self.config.get("unit_aliases", DEFAULT_APP_CONFIG["unit_aliases"])
        for a_k, a_v in curr_aliases.items():
            self.tree_alias.insert("", "end", values=(a_k, a_v))

        btn_box = tk.Frame(tab3, bg="#FFFFFF", pady=6)
        btn_box.pack(fill=tk.X)

        tk.Button(btn_box, text="➕ Thêm Từ Khóa", font=("Segoe UI", 8, "bold"), bg="#0284C7", fg="white", padx=10, pady=3, command=self._on_add_alias).pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(btn_box, text="🗑️ Xóa Đã Chọn", font=("Segoe UI", 8), bg="#EF4444", fg="white", padx=10, pady=3, command=self._on_del_alias).pack(side=tk.LEFT)

        # Footer Buttons
        footer = tk.Frame(self, bg="#07172C", padx=16, pady=10)
        footer.pack(fill=tk.X, side=tk.BOTTOM)

        btn_save = tk.Button(footer, text="💾 Lưu Cài Đặt (Áp dụng ngay)", font=("Segoe UI", 9, "bold"), bg="#15803D", fg="white", padx=16, pady=6, cursor="hand2", command=self._on_save)
        btn_save.pack(side=tk.LEFT, padx=(0, 8))

        btn_reset = tk.Button(footer, text="🔄 Khôi Phục Mặc Định", font=("Segoe UI", 9), bg="#475569", fg="white", padx=12, pady=6, cursor="hand2", command=self._on_reset)
        btn_reset.pack(side=tk.LEFT, padx=(0, 8))

        btn_close = tk.Button(footer, text="Đóng", font=("Segoe UI", 9), bg="#334155", fg="white", padx=14, pady=6, cursor="hand2", command=self.destroy)
        btn_close.pack(side=tk.RIGHT)

    def _on_add_alias(self):
        dlg = tk.Toplevel(self)
        dlg.title("Thêm Quy Tắc Viết Tắt")
        dlg.geometry("450x180")
        dlg.resizable(False, False)
        dlg.configure(bg="#F8FAFC")
        dlg.transient(self)

        tk.Label(dlg, text="Từ khóa viết tắt (Ví dụ: P2, KV4, P10):", font=("Segoe UI", 9, "bold"), bg="#F8FAFC").pack(anchor="w", padx=14, pady=(12, 2))
        ent_k = ttk.Entry(dlg, width=45)
        ent_k.pack(anchor="w", padx=14, pady=(0, 8))

        tk.Label(dlg, text="Tên đơn vị chuẩn hóa đầy đủ:", font=("Segoe UI", 9, "bold"), bg="#F8FAFC").pack(anchor="w", padx=14, pady=(2, 2))
        ent_v = ttk.Entry(dlg, width=45)
        ent_v.pack(anchor="w", padx=14, pady=(0, 10))

        def _do_add():
            k = ent_k.get().strip()
            v = ent_v.get().strip()
            if not k or not v:
                messagebox.showwarning("Thiếu dữ liệu", "Vui lòng nhập đầy đủ từ khóa và tên đơn vị.", parent=dlg)
                return
            self.tree_alias.insert("", "end", values=(k, v))
            dlg.destroy()

        btn_f = tk.Frame(dlg, bg="#F8FAFC")
        btn_f.pack(fill=tk.X, padx=14)
        tk.Button(btn_f, text="Lưu Quy Tắc", font=("Segoe UI", 9, "bold"), bg="#15803D", fg="white", padx=10, pady=4, command=_do_add).pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(btn_f, text="Hủy", font=("Segoe UI", 9), bg="#64748B", fg="white", padx=10, pady=4, command=dlg.destroy).pack(side=tk.LEFT)

    def _on_del_alias(self):
        sel = self.tree_alias.selection()
        if not sel:
            messagebox.showinfo("Chọn dòng", "Vui lòng chọn dòng quy tắc cần xóa.", parent=self)
            return
        for s in sel:
            self.tree_alias.delete(s)

    def _on_reset(self):
        if messagebox.askyesno("Khôi phục mặc định", "Anh có chắc chắn muốn khôi phục lại toàn bộ cài đặt gốc ban đầu?", parent=self):
            self.ent_agency.delete(0, tk.END)
            self.ent_agency.insert(0, DEFAULT_APP_CONFIG["agency_name"])
            self.ent_author.delete(0, tk.END)
            self.ent_author.insert(0, DEFAULT_APP_CONFIG["author_title"])
            self.sp_urgent.set(30)
            self.sp_vurgent.set(60)
            self.sp_warning.set(90)
            for item in self.tree_alias.get_children():
                self.tree_alias.delete(item)
            for k, v in DEFAULT_APP_CONFIG["unit_aliases"].items():
                self.tree_alias.insert("", "end", values=(k, v))

    def _on_save(self):
        try:
            urg = int(self.sp_urgent.get().strip() or 30)
            vurg = int(self.sp_vurgent.get().strip() or 60)
            war = int(self.sp_warning.get().strip() or 90)
        except ValueError:
            messagebox.showerror("Lỗi dữ liệu", "Ngưỡng ngày cảnh báo phải là số nguyên dương.", parent=self)
            return

        aliases = {}
        for item in self.tree_alias.get_children():
            vals = self.tree_alias.item(item, "values")
            if vals and len(vals) >= 2:
                aliases[str(vals[0]).strip()] = str(vals[1]).strip()

        new_cfg = {
            "agency_name": self.ent_agency.get().strip() or DEFAULT_APP_CONFIG["agency_name"],
            "author_title": self.ent_author.get().strip() or DEFAULT_APP_CONFIG["author_title"],
            "warning_thresholds": {
                "urgent_days": urg,
                "very_urgent_days": vurg,
                "warning_days": war
            },
            "unit_aliases": aliases
        }

        if save_app_config(new_cfg):
            global CURRENT_CONFIG
            CURRENT_CONFIG = new_cfg
            if self.on_save_cb:
                self.on_save_cb(new_cfg)
            messagebox.showinfo("Thành công", "Đã lưu cài đặt cấu hình thành công!\nCác quy tắc mới đã được áp dụng ngay lập tức.", parent=self)
            self.destroy()
        else:
            messagebox.showerror("Lỗi", "Không thể lưu tệp cấu hình config_qlahs.json.", parent=self)


class ReportSummaryDialog(tk.Toplevel):
    """
    Cửa sổ Báo cáo trích yếu số liệu nghiệp vụ:
    - Hiển thị văn bản trích yếu chuẩn phong cách ngành Kiểm sát
    - 1-Click Copy toàn bộ nội dung vào Clipboard
    - Lưu tệp văn bản Word (.doc) hoặc Text (.txt)
    """
    def __init__(self, parent, records, data_format="CASE", config=None):
        super().__init__(parent)
        self.parent = parent
        self.records = records or []
        self.data_format = data_format
        self.config = config or CURRENT_CONFIG

        self.title("📝 BÁO CÁO TRÍCH YẾU TỔNG HỢP NGHIỆP VỤ & THỜI HIỆU TRUY CỨU TNHS")
        self.geometry("920x680")
        self.minsize(800, 550)
        self.transient(parent)
        self.configure(bg="#F8FAFC")

        icon_path = get_resource_path("logo_vks.ico")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        self._build_ui()
        self.grab_set()
        self.focus_set()

    def _build_ui(self):
        # Header
        header = tk.Frame(self, bg="#07172C", padx=16, pady=10)
        header.pack(fill=tk.X)
        tk.Label(header, text="📝 BÁO CÁO TRÍCH YẾU NGHIỆP VỤ & TIẾN ĐỘ THỜI HIỆU TRUY CỨU TNHS", font=("Segoe UI", 12, "bold"), fg="#FBBF24", bg="#07172C").pack(anchor="w")
        tk.Label(header, text="Đoạn văn bản trích yếu số liệu chuẩn hóa, sẵn sàng sao chép nhanh (Ctrl+C) hoặc xuất tệp gửi Lãnh đạo Viện / Lãnh đạo Phòng.", font=("Segoe UI", 8, "italic"), fg="#93C5FD", bg="#07172C").pack(anchor="w")

        # Toolbar
        tb = tk.Frame(self, bg="#FFFFFF", padx=14, pady=8, relief=tk.RIDGE, bd=1)
        tb.pack(fill=tk.X, padx=12, pady=(8, 4))

        btn_copy = tk.Button(tb, text="📋 Sao Chép Toàn Bộ Trích Yếu (Copy)", font=("Segoe UI", 9, "bold"), bg="#0284C7", fg="white", padx=14, pady=5, cursor="hand2", command=self._on_copy)
        btn_copy.pack(side=tk.LEFT, padx=(0, 8))

        btn_txt = tk.Button(tb, text="💾 Lưu File Text (.txt)", font=("Segoe UI", 9), bg="#15803D", fg="white", padx=12, pady=5, cursor="hand2", command=self._on_save_txt)
        btn_txt.pack(side=tk.LEFT, padx=(0, 8))

        btn_doc = tk.Button(tb, text="📄 Lưu File Word (.doc)", font=("Segoe UI", 9), bg="#4338CA", fg="white", padx=12, pady=5, cursor="hand2", command=self._on_save_doc)
        btn_doc.pack(side=tk.LEFT, padx=(0, 8))

        self.lbl_copy_toast = tk.Label(tb, text="", font=("Segoe UI", 9, "bold"), fg="#16A34A", bg="#FFFFFF")
        self.lbl_copy_toast.pack(side=tk.LEFT, padx=10)

        # Main Text Widget
        text_frame = tk.Frame(self, bg="#F8FAFC", padx=12, pady=4)
        text_frame.pack(fill=tk.BOTH, expand=True)

        self.txt_content = scrolledtext.ScrolledText(text_frame, font=("Segoe UI", 10), bg="#FFFFFF", fg="#0F172A", wrap=tk.WORD, relief=tk.SOLID, bd=1, padx=12, pady=12)
        self.txt_content.pack(fill=tk.BOTH, expand=True)

        report_text = generate_executive_brief_text(self.records, self.data_format, self.config)
        self.txt_content.insert(tk.END, report_text)

        # Footer
        footer = tk.Frame(self, bg="#07172C", padx=16, pady=8)
        footer.pack(fill=tk.X, side=tk.BOTTOM)
        tk.Label(footer, text="⚡ Phát triển bởi: Trần Văn Khoa - Kiểm tra viên / VKSND TP.HCM", font=("Segoe UI", 8, "italic"), fg="#94A3B8", bg="#07172C").pack(side=tk.LEFT)
        tk.Button(footer, text="Đóng", font=("Segoe UI", 9), bg="#475569", fg="white", padx=14, pady=4, cursor="hand2", command=self.destroy).pack(side=tk.RIGHT)

    def _on_copy(self):
        text = self.txt_content.get("1.0", tk.END).strip()
        self.clipboard_clear()
        self.clipboard_append(text)
        self.lbl_copy_toast.config(text="✔ Đã sao chép vào bộ nhớ đệm (Clipboard)!")
        self.after(3500, lambda: self.lbl_copy_toast.config(text=""))

    def _on_save_txt(self):
        f = filedialog.asksaveasfilename(
            parent=self,
            title="Lưu Báo Cáo Trích Yếu",
            defaultextension=".txt",
            filetypes=[("Text File", "*.txt"), ("All Files", "*.*")],
            initialfile=f"Trich_Yeu_Bao_Cao_QLAHS_{datetime.now().strftime('%d%m%Y_%H%h%M')}.txt"
        )
        if f:
            try:
                with open(f, "w", encoding="utf-8") as fp:
                    fp.write(self.txt_content.get("1.0", tk.END))
                messagebox.showinfo("Thành công", f"Đã lưu tệp báo cáo trích yếu tại:\n{f}", parent=self)
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể lưu tệp: {e}", parent=self)

    def _on_save_doc(self):
        f = filedialog.asksaveasfilename(
            parent=self,
            title="Lưu Báo Cáo Trích Yếu Word",
            defaultextension=".doc",
            filetypes=[("Word Document", "*.doc"), ("All Files", "*.*")],
            initialfile=f"Trich_Yeu_Bao_Cao_QLAHS_{datetime.now().strftime('%d%m%Y_%H%h%M')}.doc"
        )
        if f:
            try:
                body = self.txt_content.get("1.0", tk.END).replace("\n", "<br/>\n")
                doc_html = f"""<html>
<head><meta charset="utf-8"><title>Trích yếu Báo cáo</title>
<style>body {{ font-family: 'Times New Roman', serif; font-size: 13pt; line-height: 1.4; margin: 2cm; }}</style>
</head>
<body>{body}</body></html>"""
                with open(f, "w", encoding="utf-8") as fp:
                    fp.write(doc_html)
                messagebox.showinfo("Thành công", f"Đã lưu tệp văn bản Word tại:\n{f}", parent=self)
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể lưu tệp: {e}", parent=self)


class DiffAnalysisDialog(tk.Toplevel):
    """
    Cửa sổ Phân tích so sánh biến động giữa 2 kỳ báo cáo QLAHS:
    - Cho phép chọn File Kỳ Trước và File Kỳ Này
    - Tự động so khớp theo Mã vụ án/tin báo và Số QĐ/Số thụ lý
    - Phân tích 4 nhóm: Mới thụ lý, Đã giải quyết, Chuyển cấp khẩn cấp, Tiếp tục tồn đọng
    - Bảng trực quan + Xuất Excel phân tích đa Sheet có biểu đồ BarChart
    """
    def __init__(self, parent, current_records=None, current_format="CASE"):
        super().__init__(parent)
        self.parent = parent
        self.current_records = current_records or []
        self.current_format = current_format

        self.records_old = []
        self.records_new = list(self.current_records)
        self.name_old = ""
        self.name_new = "Dữ liệu hiện tại trên màn hình" if self.current_records else ""
        self.diff_res = None

        self.title("🔄 SO SÁNH BIẾN ĐỘNG DỮ LIỆU GIỮA 2 KỲ BÁO CÁO QLAHS")
        self.geometry("1100x720")
        self.minsize(950, 600)
        self.transient(parent)
        self.configure(bg="#F8FAFC")

        icon_path = get_resource_path("logo_vks.ico")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        self._build_ui()
        self.grab_set()
        self.focus_set()

    def _build_ui(self):
        # 1. Header Banner
        header = tk.Frame(self, bg="#07172C", padx=16, pady=10)
        header.pack(fill=tk.X)
        tk.Label(header, text="🏛️  VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH", font=("Segoe UI", 9, "bold"), fg="#FBBF24", bg="#07172C").pack(anchor="w")
        tk.Label(header, text="🔄 PHÂN TÍCH SO SÁNH BIẾN ĐỘNG DỮ LIỆU GIỮA 2 KỲ BÁO CÁO QLAHS", font=("Segoe UI", 13, "bold"), fg="#FFFFFF", bg="#07172C").pack(anchor="w", pady=(2, 0))

        # 2. File Selection Bar
        sel_box = tk.Frame(self, bg="#FFFFFF", padx=16, pady=12, relief=tk.RIDGE, bd=1)
        sel_box.pack(fill=tk.X, padx=12, pady=(8, 4))

        # Row 1: File Kỳ 1 (Cũ)
        r1 = tk.Frame(sel_box, bg="#FFFFFF")
        r1.pack(fill=tk.X, pady=(0, 6))
        tk.Button(r1, text="📂 1. Chọn File Kỳ Trước (Cũ)", font=("Segoe UI", 9, "bold"), bg="#1D4ED8", fg="white", padx=12, pady=4, cursor="hand2", command=self._on_choose_old).pack(side=tk.LEFT, padx=(0, 10))
        self.lbl_old = tk.Label(r1, text="Chưa chọn tệp dữ liệu kỳ trước...", font=("Segoe UI", 9, "italic"), fg="#64748B", bg="#FFFFFF")
        self.lbl_old.pack(side=tk.LEFT)

        # Row 2: File Kỳ 2 (Mới)
        r2 = tk.Frame(sel_box, bg="#FFFFFF")
        r2.pack(fill=tk.X, pady=(0, 8))
        tk.Button(r2, text="📂 2. Chọn File Kỳ Này (Mới)", font=("Segoe UI", 9, "bold"), bg="#0D9488", fg="white", padx=12, pady=4, cursor="hand2", command=self._on_choose_new).pack(side=tk.LEFT, padx=(0, 10))
        txt_new = f"{self.name_new} ({len(self.records_new):,} hồ sơ)" if self.records_new else "Chưa chọn tệp dữ liệu kỳ này..."
        self.lbl_new = tk.Label(r2, text=txt_new, font=("Segoe UI", 9, "bold" if self.records_new else "italic"), fg="#047857" if self.records_new else "#64748B", bg="#FFFFFF")
        self.lbl_new.pack(side=tk.LEFT)

        # Row 3: Nút Chạy đối chiếu
        r3 = tk.Frame(sel_box, bg="#FFFFFF")
        r3.pack(fill=tk.X)
        btn_run = tk.Button(r3, text="⚡ TIẾN HÀNH PHÂN TÍCH SO SÁNH BIẾN ĐỘNG", font=("Segoe UI", 10, "bold"), bg="#7C3AED", fg="white", activebackground="#6D28D9", activeforeground="white", padx=16, pady=6, cursor="hand2", command=self._on_run_diff)
        btn_run.pack(side=tk.LEFT, padx=(0, 10))
        self.lbl_diff_status = tk.Label(r3, text="Sẵn sàng nạp 2 tệp để so sánh biến động.", font=("Segoe UI", 9, "italic"), fg="#64748B", bg="#FFFFFF")
        self.lbl_diff_status.pack(side=tk.LEFT)

        # 3. KPI Cards Frame
        self.kpi_frame = tk.Frame(self, bg="#F8FAFC", padx=12, pady=4)
        self.kpi_frame.pack(fill=tk.X)

        self.kpi_labels = {}
        cards_cfg = [
            ("added", "✨ MỚI THỤ LÝ TRONG KỲ", "0", "Phát sinh mới", "#047857", "#ECFDF5", "#A7F3D0"),
            ("resolved", "✅ ĐÃ GIẢI QUYẾT XONG", "0", "Đạt kết quả trong kỳ", "#1D4ED8", "#EFF6FF", "#BFDBFE"),
            ("worsened", "⚠️ NGUY CƠ THỜI HIỆU MỚI", "0", "Chuyển cấp khẩn cấp", "#B91C1C", "#FEF2F2", "#FECACA"),
            ("continued", "⏸ TỒN ĐỌNG TIẾP TỤC", "0", "Đang tiếp tục thụ lý", "#475569", "#F8FAFC", "#E2E8F0"),
        ]

        for idx, (cid, title, val, sub, fg_c, bg_c, bd_c) in enumerate(cards_cfg):
            card = tk.Frame(self.kpi_frame, bg=bg_c, highlightthickness=1, highlightbackground=bd_c, padx=10, pady=6)
            card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4)
            tk.Label(card, text=title, font=("Segoe UI", 8, "bold"), fg=fg_c, bg=bg_c).pack(anchor="w")
            lbl_v = tk.Label(card, text=val, font=("Segoe UI", 16, "bold"), fg=fg_c, bg=bg_c)
            lbl_v.pack(anchor="w")
            tk.Label(card, text=sub, font=("Segoe UI", 7, "italic"), fg="#64748B", bg=bg_c).pack(anchor="w")
            self.kpi_labels[cid] = lbl_v

        # 4. Canvas Thanh Biểu đồ trực quan
        self.chart_canvas = tk.Canvas(self, bg="#FFFFFF", height=38, highlightthickness=1, highlightbackground="#CBD5E1")
        self.chart_canvas.pack(fill=tk.X, padx=16, pady=(4, 6))

        # 5. Notebook Tabs hiển thị chi tiết
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill=tk.BOTH, expand=True, padx=12, pady=4)

        self.tree_added = self._create_tab_tree("tab_added", "🟢 Mới Thụ Lý (0)")
        self.tree_resolved = self._create_tab_tree("tab_resolved", "🔵 Đã Giải Quyết Xong (0)")
        self.tree_worsened = self._create_tab_tree("tab_worsened", "🔴 Chuyển Cấp Khẩn Cấp (0)")
        self.tree_continued = self._create_tab_tree("tab_continued", "⚪ Tồn Đọng Tiếp Tục (0)")

        # 6. Action Bar (Bottom)
        act_bar = tk.Frame(self, bg="#07172C", padx=16, pady=8)
        act_bar.pack(fill=tk.X, side=tk.BOTTOM)

        self.btn_export_diff = tk.Button(act_bar, text="📥 Xuất Báo Cáo Biến Động Ra Excel (.xlsx)", font=("Segoe UI", 9, "bold"), bg="#15803D", fg="white", padx=14, pady=5, cursor="hand2", command=self._on_export_diff)
        self.btn_export_diff.pack(side=tk.LEFT, padx=(0, 8))

        btn_brief = tk.Button(act_bar, text="📝 Trích Yếu Số Liệu", font=("Segoe UI", 9), bg="#0284C7", fg="white", padx=12, pady=5, cursor="hand2", command=self._on_show_brief)
        btn_brief.pack(side=tk.LEFT, padx=(0, 8))

        btn_close = tk.Button(act_bar, text="Đóng", font=("Segoe UI", 9), bg="#475569", fg="white", padx=14, pady=5, cursor="hand2", command=self.destroy)
        btn_close.pack(side=tk.RIGHT)

    def _create_tab_tree(self, tab_id, tab_title):
        tab_frame = tk.Frame(self.nb, bg="#FFFFFF")
        self.nb.add(tab_frame, text=tab_title)

        cols = ("stt", "code", "title", "unit", "ksv", "due", "status")
        tree = ttk.Treeview(tab_frame, columns=cols, show="headings")
        tree.heading("stt", text="STT")
        tree.heading("code", text="Mã số hồ sơ")
        tree.heading("title", text="Tên vụ án / Nội dung tin báo")
        tree.heading("unit", text="Đơn vị VKS thụ lý")
        tree.heading("ksv", text="Kiểm sát viên")
        tree.heading("due", text="Hạn / Thời hiệu")
        tree.heading("status", text="Trạng thái / Đánh giá")

        tree.column("stt", width=45, anchor=tk.CENTER)
        tree.column("code", width=140, anchor=tk.CENTER)
        tree.column("title", width=340, anchor=tk.W)
        tree.column("unit", width=180, anchor=tk.W)
        tree.column("ksv", width=140, anchor=tk.W)
        tree.column("due", width=110, anchor=tk.CENTER)
        tree.column("status", width=160, anchor=tk.W)

        tree.tag_configure("odd", background="#F8FAFC")
        tree.tag_configure("even", background="#FFFFFF")

        sb = ttk.Scrollbar(tab_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        return tree

    def _on_choose_old(self):
        f = filedialog.askopenfilename(parent=self, title="Chọn File QLAHS Kỳ Trước (Cũ)", filetypes=[("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")])
        if f:
            try:
                wb = openpyxl.load_workbook(f, data_only=True)
                raw_rows = list(wb.active.iter_rows(values_only=True))
                fmt = detect_data_format(raw_rows)
                if fmt == "TIN_BAO":
                    recs = parse_qlahs_tin_bao_rows(raw_rows)
                elif fmt == "CASE_REPORT":
                    recs = parse_qlahs_case_report_rows(raw_rows, filename=os.path.basename(f))
                else:
                    recs = parse_qlahs_case_rows(raw_rows)
                self.records_old = recs
                self.name_old = os.path.basename(f)
                self.lbl_old.config(text=f"{self.name_old} ({len(recs):,} hồ sơ)", fg="#1D4ED8", font=("Segoe UI", 9, "bold"))
            except Exception as e:
                messagebox.showerror("Lỗi đọc file", f"Không thể đọc file kỳ trước: {e}", parent=self)

    def _on_choose_new(self):
        f = filedialog.askopenfilename(parent=self, title="Chọn File QLAHS Kỳ Này (Mới)", filetypes=[("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")])
        if f:
            try:
                wb = openpyxl.load_workbook(f, data_only=True)
                raw_rows = list(wb.active.iter_rows(values_only=True))
                fmt = detect_data_format(raw_rows)
                if fmt == "TIN_BAO":
                    recs = parse_qlahs_tin_bao_rows(raw_rows)
                elif fmt == "CASE_REPORT":
                    recs = parse_qlahs_case_report_rows(raw_rows, filename=os.path.basename(f))
                else:
                    recs = parse_qlahs_case_rows(raw_rows)
                self.records_new = recs
                self.name_new = os.path.basename(f)
                self.lbl_new.config(text=f"{self.name_new} ({len(recs):,} hồ sơ)", fg="#047857", font=("Segoe UI", 9, "bold"))
            except Exception as e:
                messagebox.showerror("Lỗi đọc file", f"Không thể đọc file kỳ này: {e}", parent=self)

    def _on_run_diff(self):
        if not self.records_old:
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng chọn File Kỳ Trước (Cũ) trước khi tiến hành phân tích.", parent=self)
            return
        if not self.records_new:
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng chọn File Kỳ Này (Mới) trước khi tiến hành phân tích.", parent=self)
            return

        self.diff_res = compare_two_qlahs_periods(self.records_old, self.records_new, self.current_format)
        st = self.diff_res["stats"]

        # Cập nhật KPI
        self.kpi_labels["added"].config(text=f"+{st['n_added']:,}")
        self.kpi_labels["resolved"].config(text=f"-{st['n_resolved']:,}")
        self.kpi_labels["worsened"].config(text=f"{st['n_worsened']:,}")
        self.kpi_labels["continued"].config(text=f"{st['n_continued']:,}")

        # Vẽ thanh biểu đồ trực quan
        w = self.chart_canvas.winfo_width() or 800
        draw_visual_distribution_chart(self.chart_canvas, w, 38, [
            ("Mới thụ lý", st["n_added"], "#10B981"),
            ("Đã giải quyết", st["n_resolved"], "#3B82F6"),
            ("Khẩn cấp mới", st["n_worsened"], "#EF4444"),
            ("Tồn đọng", st["n_continued"], "#94A3B8"),
        ])

        # Điền dữ liệu vào các Treeview
        def populate_tree(tree, items, is_worsened=False):
            for itm in tree.get_children():
                tree.delete(itm)
            for idx, r in enumerate(items, start=1):
                code = r.get("ma_tin_bao") or r.get("ma_vu_an") or ""
                title = r.get("noi_dung_tin_bao") or r.get("ten_vu_an") or ""
                unit = r.get("vks_thu_ly") or ""
                ksv = r.get("ksv_chinh") or r.get("can_bo_giai_quyet") or ""
                due = r.get("thoi_han_den_ngay") or r.get("ngay_het_thoi_hieu") or ""
                if is_worsened:
                    status = f"Từ {r.get('_old_urgency')} -> {r.get('_new_urgency')}"
                else:
                    status = r.get("ket_qua_giai_quyet") or r.get("muc_do_khan_cap") or r.get("canh_bao_thoi_hieu") or "Đang thụ lý"
                tag = "even" if idx % 2 == 0 else "odd"
                tree.insert("", "end", values=(idx, code, title, unit, ksv, due, status), tags=(tag,))

        populate_tree(self.tree_added, self.diff_res["added"])
        populate_tree(self.tree_resolved, self.diff_res["resolved"])
        populate_tree(self.tree_worsened, self.diff_res["worsened"], is_worsened=True)
        populate_tree(self.tree_continued, self.diff_res["continued"][:500])

        self.nb.tab(0, text=f"🟢 Mới Thụ Lý ({st['n_added']:,})")
        self.nb.tab(1, text=f"🔵 Đã Giải Quyết ({st['n_resolved']:,})")
        self.nb.tab(2, text=f"🔴 Chuyển Cấp Khẩn Cấp ({st['n_worsened']:,})")
        self.nb.tab(3, text=f"⚪ Tồn Đọng ({st['n_continued']:,})")

        self.lbl_diff_status.config(
            text=f"✔ Hoàn tất phân tích! Chênh lệch ròng: {'+' if st['net_change'] >= 0 else ''}{st['net_change']} hồ sơ.",
            fg="#16A34A",
            font=("Segoe UI", 9, "bold")
        )

    def _on_export_diff(self):
        if not self.diff_res:
            messagebox.showwarning("Chưa có kết quả", "Vui lòng bấm 'Tiến Hành Phân Tích' trước khi xuất file.", parent=self)
            return

        now_str = datetime.now().strftime("%d%m%Y_%H%h%M")
        default_fn = f"Bao_Cao_Bien_Dong_QLAHS_{now_str}.xlsx"
        out_f = filedialog.asksaveasfilename(
            parent=self,
            title="Lưu Báo Cáo Biến Động Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            initialfile=default_fn
        )
        if out_f:
            try:
                export_diff_to_excel(self.diff_res, out_f, self.name_old or "Kỳ Trước", self.name_new or "Kỳ Này")
                messagebox.showinfo("Thành công", f"Đã xuất file báo cáo so sánh biến động Excel thành công tại:\n{out_f}", parent=self)
            except Exception as e:
                messagebox.showerror("Lỗi xuất file", f"Không thể lưu file Excel: {e}", parent=self)

    def _on_show_brief(self):
        if not self.diff_res:
            messagebox.showwarning("Chưa có kết quả", "Vui lòng bấm 'Tiến Hành Phân Tích' trước.", parent=self)
            return
        st = self.diff_res["stats"]
        txt = f"""VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH
TRÍCH YẾU SO SÁNH BIẾN ĐỘNG GIỮA 2 KỲ BÁO CÁO QLAHS
--------------------------------------------------------------------------------
• Kỳ trước: {self.name_old} ({st['total_old']:,} hồ sơ)
• Kỳ này:   {self.name_new} ({st['total_new']:,} hồ sơ)
• Biến động ròng: {'+' if st['net_change'] >= 0 else ''}{st['net_change']:,} hồ sơ

1. HỒ SƠ PHÁT SINH MỚI: +{st['n_added']:,} hồ sơ mới thụ lý trong kỳ.
2. HỒ SƠ ĐÃ GIẢI QUYẾT: -{st['n_resolved']:,} hồ sơ đã giải quyết / kết thúc.
3. CẢNH BÁO NGUY CƠ:    {st['n_worsened']:,} hồ sơ chuyển mức khẩn cấp thời hiệu TNHS.
4. TỒN ĐỌNG TIẾP TỤC:   {st['n_continued']:,} hồ sơ tiếp tục theo dõi giải quyết.
--------------------------------------------------------------------------------
Tác giả: {CURRENT_CONFIG.get('author_title')}"""
        dlg_b = tk.Toplevel(self)
        dlg_b.title("Trích Yếu Biến Động")
        dlg_b.geometry("600x380")
        dlg_b.configure(bg="#F8FAFC")
        dlg_b.transient(self)

        t_area = scrolledtext.ScrolledText(dlg_b, font=("Segoe UI", 10), wrap=tk.WORD, padx=10, pady=10)
        t_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        t_area.insert(tk.END, txt)

        def _cp():
            self.clipboard_clear()
            self.clipboard_append(txt)
            messagebox.showinfo("Đã sao chép", "Đã chép trích yếu vào Clipboard!", parent=dlg_b)

        tk.Button(dlg_b, text="📋 Sao Chép (Copy)", font=("Segoe UI", 9, "bold"), bg="#0284C7", fg="white", padx=12, pady=4, command=_cp).pack(pady=(0, 10))


class QLAHSApp(tk.Tk):
    def __init__(self):
        super().__init__()
        # Cập nhật tiêu đề theo yêu cầu
        self.title("VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH - Hệ thống Chuẩn hóa & Đối chiếu QLAHS (Version 1.0 Offline)")
        self.geometry("1400x820")
        self.minsize(1100, 680)
        self.configure(bg="#F8FAFC")

        # Cài đặt biểu tượng cửa sổ nếu có file logo_vks.ico
        icon_path = get_resource_path("logo_vks.ico")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        self.data_format = "DEFENDANT"
        self.records = []
        self._search_job = None
        self.kpi_titles = {}
        self.kpi_cards = {}
        self._setup_styles()
        self._build_ui()

    def run_async_task(self, title, loading_msg, worker_fn, on_success, on_error=None):
        """Chạy tác vụ nặng ở luồng phụ (Thread) an toàn tuyệt đối qua Queue với Loading Bar."""
        dlg = LoadingDialog(self, title=str(title), message=str(loading_msg))
        dlg.update()
        self.update()

        task_queue = queue.Queue()

        def _worker_thread():
            try:
                def progress_cb(msg, sub=None):
                    task_queue.put(("progress", msg, sub))

                result = worker_fn(progress_cb)
                task_queue.put(("done", result))
            except Exception as e:
                task_queue.put(("error", e))

        def _check_queue():
            try:
                while True:
                    item = task_queue.get_nowait()
                    msg_type = item[0]
                    if msg_type == "progress":
                        dlg.update_message(item[1], item[2])
                    elif msg_type == "done":
                        try:
                            dlg.close()
                        except Exception:
                            pass
                        self.update_idletasks()
                        try:
                            on_success(item[1])
                        except Exception as ex:
                            messagebox.showerror("Lỗi hiển thị", f"Không thể cập nhật giao diện:\n{ex}")
                        return
                    elif msg_type == "error":
                        dlg.close()
                        err = item[1]
                        if on_error:
                            on_error(err)
                        else:
                            messagebox.showerror("Thông báo lỗi", f"Quá trình xử lý dữ liệu gặp lỗi:\n{err}")
                        return
            except queue.Empty:
                pass

            self.after(40, _check_queue)

        threading.Thread(target=_worker_thread, daemon=True).start()
        self.after(40, _check_queue)

    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        # Style Treeview Headings: Deep Navy với chữ trắng nổi bật
        style.configure(
            "Treeview.Heading",
            background="#0A2540",
            foreground="#FFFFFF",
            font=("Segoe UI", 9, "bold"),
            relief=tk.FLAT,
            padding=[4, 6]
        )
        style.map(
            "Treeview.Heading",
            background=[("active", "#153E75"), ("pressed", "#991B1B")]
        )

        # Style Notebook Tabs: Phong cách VKS sinh động
        style.configure("TNotebook", background="#F8FAFC", borderwidth=0)
        style.configure(
            "TNotebook.Tab",
            font=("Segoe UI", 10, "bold"),
            padding=[20, 8],
            background="#E2E8F0",
            foreground="#334155"
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", "#0A2540"), ("active", "#CBD5E1")],
            foreground=[("selected", "#FBBF24"), ("active", "#0F172A")]
        )

        # Style Treeview Rows
        style.configure("Treeview", font=("Segoe UI", 9), rowheight=26, borderwidth=1, relief=tk.SOLID)
        style.map("Treeview", background=[("selected", "#DBEAFE")], foreground=[("selected", "#1E3A8A")])

    def _build_fallback_logo(self, parent):
        box = tk.Frame(parent, bg="#991B1B", width=48, height=48, relief=tk.RAISED, bd=2)
        box.pack(side=tk.LEFT, padx=(0, 12))
        box.pack_propagate(False)
        tk.Label(box, text="⚖", font=("Segoe UI", 20, "bold"), fg="#FEF08A", bg="#991B1B").pack(expand=True)

    def _build_ui(self):
        # 1. Header chủ đề VKSND TP. HỒ CHÍ MINH sinh động
        header = tk.Frame(self, bg="#07172C", height=76)
        header.pack(fill=tk.X)

        # Vạch màu vàng cờ / hoàng kim trên cùng
        top_stripe = tk.Frame(header, bg="#F59E0B", height=4)
        top_stripe.pack(fill=tk.X, side=tk.TOP)

        header_content = tk.Frame(header, bg="#07172C")
        header_content.pack(fill=tk.BOTH, expand=True, padx=16, pady=6)

        # Logo / Hình ảnh Trụ sở VKSND TP. Hồ Chí Minh
        logo_img_path = get_resource_path("logo_header_vksnd_tphcm.png")
        if os.path.exists(logo_img_path):
            try:
                self.header_logo_img = tk.PhotoImage(file=logo_img_path)
                logo_lbl = tk.Label(
                    header_content,
                    image=self.header_logo_img,
                    bg="#07172C",
                    relief=tk.SOLID,
                    bd=1,
                    highlightbackground="#F59E0B",
                    highlightthickness=2
                )
                logo_lbl.pack(side=tk.LEFT, padx=(0, 12))
            except Exception:
                self._build_fallback_logo(header_content)
        else:
            self._build_fallback_logo(header_content)

        # Khối tiêu đề
        title_box = tk.Frame(header_content, bg="#07172C")
        title_box.pack(side=tk.LEFT, fill=tk.Y)

        tk.Label(
            title_box,
            text="VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH",
            font=("Segoe UI", 11, "bold"),
            fg="#FBBF24",
            bg="#07172C",
        ).pack(anchor=tk.W)

        tk.Label(
            title_box,
            text="HỆ THỐNG TỔNG HỢP & XỬ LÝ DỮ LIỆU QLAHS (VERSION 1.0)",
            font=("Segoe UI", 14, "bold"),
            fg="#FFFFFF",
            bg="#07172C",
        ).pack(anchor=tk.W)

        tk.Label(
            title_box,
            text="Tác giả: Trần Văn Khoa - Kiểm tra viên / VKSND TP.HCM  •  Bản dùng thử (v1.0 Offline)",
            font=("Segoe UI", 8, "italic"),
            fg="#93C5FD",
            bg="#07172C",
        ).pack(anchor=tk.W, pady=(2, 0))

        # Khối huy hiệu & châm ngôn ngành bên phải
        right_box = tk.Frame(header_content, bg="#07172C")
        right_box.pack(side=tk.RIGHT, fill=tk.Y)

        tk.Label(
            right_box,
            text="BẢO MẬT OFFLINE 100%",
            font=("Segoe UI", 8, "bold"),
            fg="#6EE7B7",
            bg="#064E3B",
            padx=8,
            pady=2,
            relief=tk.FLAT
        ).pack(anchor=tk.E, pady=(0, 4))

        tk.Label(
            right_box,
            text="★ CÔNG MINH - CHÍNH TRỰC - KHÁCH QUAN - THẬN TRỌNG - KHIÊM TỐN ★",
            font=("Segoe UI", 9, "bold"),
            fg="#FEF08A",
            bg="#831843",
            padx=10,
            pady=3,
            relief=tk.GROOVE,
        ).pack(anchor=tk.E)

        # 2. THANH THẺ KPI TỔNG HỢP MÀU SẮC SINH ĐỘNG (5 Khối trực quan)
        kpi_bar = tk.Frame(self, bg="#F1F5F9", pady=6, padx=12)
        kpi_bar.pack(fill=tk.X)

        self.kpi_cards = {}
        cards_config = [
            ("total", "TỔNG SỐ ÁN QUẢN LÝ", "#EFF6FF", "#3B82F6", "#1E40AF", "0 vụ án"),
            ("cases", "⚠️ ĐÃ HẾT THỜI HIỆU (Đ.230)", "#FEF2F2", "#EF4444", "#991B1B", "0 vụ án"),
            ("provinces", "🔴 KHẨN CẤP (≤ 15 NGÀY)", "#FFFBEB", "#F59E0B", "#B45309", "0 vụ án"),
            ("stages", "⏸️ TẠM ĐÌNH CHỈ / ĐÌNH CHỈ", "#F5F3FF", "#8B5CF6", "#6B21A8", "0 vụ án"),
            ("status", "⚖️ PHÂN BỔ GIAI ĐOẠN", "#ECFDF5", "#10B981", "#047857", "Chưa có"),
        ]

        for key, title, bg_col, border_col, text_col, init_val in cards_config:
            c_frame = tk.Frame(kpi_bar, bg=bg_col, relief=tk.SOLID, bd=1, highlightbackground=border_col, highlightthickness=1)
            c_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4, ipady=3)

            title_lbl = tk.Label(
                c_frame,
                text=title,
                font=("Segoe UI", 8, "bold"),
                fg=text_col,
                bg=bg_col
            )
            title_lbl.pack(anchor=tk.W, padx=8, pady=(2, 0))
            self.kpi_titles[key] = title_lbl

            val_lbl = tk.Label(
                c_frame,
                text=init_val,
                font=("Segoe UI", 12 if key in ("total", "cases", "provinces") else 9, "bold"),
                fg=text_col,
                bg=bg_col
            )
            val_lbl.pack(anchor=tk.W, padx=8, pady=(0, 2))
            self.kpi_cards[key] = val_lbl

        # 3. Toolbar các nút chức năng sinh động
        toolbar = tk.Frame(self, bg="#FFFFFF", pady=8, padx=16, relief=tk.RIDGE, bd=1)
        toolbar.pack(fill=tk.X, padx=12, pady=(4, 4))

        btn_open = tk.Button(
            toolbar,
            text="📂 1. Chọn File (.xlsx)",
            font=("Segoe UI", 9, "bold"),
            bg="#1D4ED8",
            fg="white",
            activebackground="#1E40AF",
            activeforeground="white",
            padx=10,
            pady=6,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.on_open_excel,
        )
        btn_open.pack(side=tk.LEFT, padx=(0, 6))

        btn_paste = tk.Button(
            toolbar,
            text="📋 2. Dán Clipboard",
            font=("Segoe UI", 9, "bold"),
            bg="#0D9488",
            fg="white",
            activebackground="#0F766E",
            activeforeground="white",
            padx=10,
            pady=6,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.on_paste_clipboard,
        )
        btn_paste.pack(side=tk.LEFT, padx=(0, 6))

        btn_export = tk.Button(
            toolbar,
            text="📥 3. Xuất Excel & Chart",
            font=("Segoe UI", 9, "bold"),
            bg="#B91C1C",
            fg="white",
            activebackground="#991B1B",
            activeforeground="white",
            padx=10,
            pady=6,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.on_export_excel,
        )
        btn_export.pack(side=tk.LEFT, padx=(0, 6))

        btn_reconcile = tk.Button(
            toolbar,
            text="⚡ 4. Đối Chiếu Hợp Nhất",
            font=("Segoe UI", 9, "bold"),
            bg="#7C3AED",
            fg="white",
            activebackground="#6D28D9",
            activeforeground="white",
            padx=10,
            pady=6,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.on_reconcile_multiple_files,
        )
        btn_reconcile.pack(side=tk.LEFT, padx=(0, 6))

        btn_diff = tk.Button(
            toolbar,
            text="🔄 5. Biến Động 2 Kỳ (Diff)",
            font=("Segoe UI", 9, "bold"),
            bg="#2563EB",
            fg="white",
            activebackground="#1D4ED8",
            activeforeground="white",
            padx=10,
            pady=6,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.on_diff_analysis,
        )
        btn_diff.pack(side=tk.LEFT, padx=(0, 6))

        btn_summary = tk.Button(
            toolbar,
            text="📝 6. Trích Yếu Báo Cáo",
            font=("Segoe UI", 9, "bold"),
            bg="#059669",
            fg="white",
            activebackground="#047857",
            activeforeground="white",
            padx=10,
            pady=6,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.on_report_summary,
        )
        btn_summary.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_dashboard = tk.Button(
            toolbar,
            text="📈 Dashboard Đối Chiếu",
            font=("Segoe UI", 9, "bold"),
            bg="#0284C7",
            fg="white",
            activebackground="#0369A1",
            activeforeground="white",
            padx=10,
            pady=6,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.show_reconciliation_dashboard,
        )
        self.btn_dashboard.pack(side=tk.LEFT, padx=(0, 6))

        btn_pivot_toggle = tk.Button(
            toolbar,
            text="📊 Ma Trận Pivot",
            font=("Segoe UI", 9, "bold"),
            bg="#D97706",
            fg="white",
            activebackground="#B45309",
            activeforeground="white",
            padx=10,
            pady=6,
            relief=tk.FLAT,
            cursor="hand2",
            command=lambda: self.notebook.select(1),
        )
        btn_pivot_toggle.pack(side=tk.LEFT, padx=(0, 6))

        btn_clear = tk.Button(
            toolbar,
            text="Xóa danh sách",
            font=("Segoe UI", 9),
            bg="#F1F5F9",
            fg="#475569",
            activebackground="#E2E8F0",
            padx=10,
            pady=6,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.on_clear,
        )
        btn_clear.pack(side=tk.RIGHT)

        btn_settings = tk.Button(
            toolbar,
            text="⚙️ Cài Đặt Quy Tắc",
            font=("Segoe UI", 9, "bold"),
            bg="#475569",
            fg="white",
            activebackground="#334155",
            activeforeground="white",
            padx=10,
            pady=6,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.on_open_settings,
        )
        btn_settings.pack(side=tk.RIGHT, padx=(0, 6))

        self.lbl_status = tk.Label(
            toolbar,
            text="Hệ thống sẵn sàng xử lý dữ liệu hồ sơ vụ án - bị can.",
            font=("Segoe UI", 9, "italic"),
            bg="#FFFFFF",
            fg="#64748B",
        )
        self.lbl_status.pack(side=tk.LEFT, padx=8)

        # 4. Thanh bản quyền tác giả chân trang (Footer Bar) - Ghim cố định đáy ứng dụng
        footer_bar = tk.Frame(self, bg="#07172C", padx=16, pady=4, relief=tk.RIDGE, bd=1)
        footer_bar.pack(side=tk.BOTTOM, fill=tk.X)

        tk.Label(
            footer_bar,
            text="© Bản quyền phần mềm Chuẩn hóa & Đối chiếu QLAHS Offline | VKSND TP.HCM",
            font=("Segoe UI", 8),
            fg="#94A3B8",
            bg="#07172C",
        ).pack(side=tk.LEFT)

        self.lbl_footer_author = tk.Label(
            footer_bar,
            text=f"Phát triển bởi: {CURRENT_CONFIG.get('author_title', 'Trần Văn Khoa - Kiểm tra viên / VKSND TP.HCM')}",
            font=("Segoe UI", 8, "italic"),
            fg="#FBBF24",
            bg="#07172C",
        )
        self.lbl_footer_author.pack(side=tk.RIGHT)

        # 5. Notebook gồm 2 Tab chính: Tab Bảng dữ liệu chuẩn hóa và Tab Pivot Table
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=12, pady=4)

        # Tab 1: Danh sách chuẩn hóa
        self.tab_table = tk.Frame(self.notebook, bg="#FFFFFF")
        self.notebook.add(self.tab_table, text="📋 Danh Sách Hồ Sơ Vụ Án Chuẩn Hóa")
        self._build_table_tab(self.tab_table)

        # Tab 2: Bảng Pivot Table Tổng hợp Điều hành
        self.tab_pivot = tk.Frame(self.notebook, bg="#FFFFFF")
        self.notebook.add(self.tab_pivot, text="📊 Bảng Pivot Table Ma Trận Điều Hành Lãnh Đạo")
        self._build_pivot_tab(self.tab_pivot)

    def _build_table_tab(self, parent):
        # Thanh tìm kiếm & lọc nhanh màu sắc sinh động
        filter_bar = tk.Frame(parent, bg="#F0FDF4", pady=6, padx=12, relief=tk.GROOVE, bd=1)
        filter_bar.pack(fill=tk.X)

        tk.Label(filter_bar, text="🔍 Tìm kiếm:", font=("Segoe UI", 9, "bold"), fg="#166534", bg="#F0FDF4").pack(side=tk.LEFT, padx=(0, 4))
        self.search_var = tk.StringVar()
        def _on_search_change(*args):
            if self._search_job is not None:
                self.after_cancel(self._search_job)
            self._search_job = self.after(200, self.refresh_views)
        self.search_var.trace_add("write", _on_search_change)
        entry_search = tk.Entry(filter_bar, textvariable=self.search_var, width=24, font=("Segoe UI", 9))
        entry_search.pack(side=tk.LEFT, padx=(0, 14))

        tk.Label(filter_bar, text="🏛️ Tỉnh/Thành:", font=("Segoe UI", 9, "bold"), fg="#075985", bg="#F0FDF4").pack(side=tk.LEFT, padx=(0, 4))
        self.prov_filter_var = tk.StringVar(value="Tất cả")
        self.cbo_prov = ttk.Combobox(filter_bar, textvariable=self.prov_filter_var, state="readonly", width=16)
        self.cbo_prov["values"] = ["Tất cả"]
        self.cbo_prov.bind("<<ComboboxSelected>>", lambda e: self.refresh_views())
        self.cbo_prov.pack(side=tk.LEFT, padx=(0, 14))

        tk.Label(filter_bar, text="⚖️ Giai đoạn:", font=("Segoe UI", 9, "bold"), fg="#0f766e", bg="#F0FDF4").pack(side=tk.LEFT, padx=(0, 4))
        self.stage_filter_var = tk.StringVar(value="Tất cả")
        self.cbo_stage = ttk.Combobox(filter_bar, textvariable=self.stage_filter_var, state="readonly", width=13)
        self.cbo_stage["values"] = ["Tất cả", "Điều tra", "Truy tố", "Sơ thẩm", "Phúc thẩm"]
        self.cbo_stage.bind("<<ComboboxSelected>>", lambda e: self.refresh_views())
        self.cbo_stage.pack(side=tk.LEFT, padx=(0, 14))

        tk.Label(filter_bar, text="🏷️ Trạng thái:", font=("Segoe UI", 9, "bold"), fg="#9a3412", bg="#F0FDF4").pack(side=tk.LEFT, padx=(0, 4))
        self.status_filter_var = tk.StringVar(value="Tất cả")
        self.cbo_status = ttk.Combobox(filter_bar, textvariable=self.status_filter_var, state="readonly", width=13)
        self.cbo_status["values"] = ["Tất cả", "Đình nã", "Truy nã", "Tạm đình chỉ"]
        self.cbo_status.bind("<<ComboboxSelected>>", lambda e: self.refresh_views())
        self.cbo_status.pack(side=tk.LEFT)

        # Bảng Treeview 12 cột
        table_frame = tk.Frame(parent, bg="#FFFFFF")
        table_frame.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        cols = [k for k, _, _ in COLUMNS_12]
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=20)

        # Cấu hình màu sắc hàng (Dòng chẵn / lẻ và màu trạng thái)
        self.tree.tag_configure("even", background="#FFFFFF")
        self.tree.tag_configure("odd", background="#F8FAFC")
        self.tree.tag_configure("dinh_na", foreground="#B45309")
        self.tree.tag_configure("truy_na", foreground="#DC2626")
        self.tree.tag_configure("sap_het_thoi_hieu", foreground="#DC2626", font=("Segoe UI", 9, "bold"))
        self.tree.tag_configure("tag_khan_cap", background="#FEE2E2", foreground="#991B1B", font=("Segoe UI", 9, "bold"))
        self.tree.tag_configure("tag_rat_gap", background="#FEF3C7", foreground="#92400E", font=("Segoe UI", 9, "bold"))
        self.tree.tag_configure("tag_sap_het", background="#FEF9C3", foreground="#854D0E")

        for key, title, width in COLUMNS_12:
            self.tree.heading(key, text=title)
            anchor = tk.CENTER if key in ("stt", "tinh_thanh", "giai_doan", "ngay_sinh", "trang_thai") else tk.W
            self.tree.column(key, width=width, anchor=anchor)

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

    def _build_pivot_tab(self, parent):
        # Khung thanh điều khiển Pivot
        p_ctrl = tk.Frame(parent, bg="#EFF6FF", pady=8, padx=12, relief=tk.GROOVE, bd=1)
        p_ctrl.pack(fill=tk.X)

        tk.Label(
            p_ctrl,
            text="CHẾ ĐỘ XEM PIVOT:",
            font=("Segoe UI", 10, "bold"),
            fg="#1E40AF",
            bg="#EFF6FF"
        ).pack(side=tk.LEFT, padx=(0, 8))

        self.cbo_pivot_mode = ttk.Combobox(
            p_ctrl,
            state="readonly",
            values=[
                "🏛️ 1. MA TRẬN TIẾN ĐỘ: Đơn Vị VKS × Giai Đoạn Tố Tụng (Điều Tra, Truy Tố, Xét Xử)",
                "🚨 2. MA TRẬN RỦI RO: Đơn Vị VKS × Cảnh Báo Thời Hiệu TNHS & Quá Hạn (Đ.230)",
                "👑 3. BÁO CÁO TOÀN DIỆN 360°: Tổng Quan Điều Hành Lãnh Đạo Viện",
                "👤 4. PHÂN CÔNG NGHIỆP VỤ: Kiểm Sát Viên × Trạng Thái Vụ Án",
                "⚖️ 5. PHÂN LOẠI TỘI PHẠM: Đơn Vị VKS × Tính Chất Tội (Ít NT, NT, RNT, ĐBNT)"
            ],
            width=65,
            font=("Segoe UI", 9, "bold")
        )
        self.cbo_pivot_mode.current(0)
        self.cbo_pivot_mode.pack(side=tk.LEFT, padx=(0, 15))
        self.cbo_pivot_mode.bind("<<ComboboxSelected>>", lambda e: self.refresh_views())

        self.lbl_pivot_desc = tk.Label(
            p_ctrl,
            text="Tự động trích xuất và tổng hợp số liệu ma trận Pivot phục vụ báo cáo lãnh đạo Viện.",
            font=("Segoe UI", 9, "italic"),
            fg="#475569",
            bg="#EFF6FF"
        )
        self.lbl_pivot_desc.pack(side=tk.LEFT)

        # Bảng Treeview ma trận Pivot
        pivot_table_frame = tk.Frame(parent, bg="#FFFFFF")
        pivot_table_frame.pack(fill=tk.BOTH, expand=True, padx=4, pady=6)

        self.pivot_tree = ttk.Treeview(pivot_table_frame, show="headings", height=18)
        self.pivot_tree.tag_configure("p_even", background="#FFFFFF")
        self.pivot_tree.tag_configure("p_odd", background="#F0FDF4")
        self.pivot_tree.tag_configure("p_total", background="#FEF3C7", font=("Segoe UI", 9, "bold"))

        p_vsb = ttk.Scrollbar(pivot_table_frame, orient="vertical", command=self.pivot_tree.yview)
        p_hsb = ttk.Scrollbar(pivot_table_frame, orient="horizontal", command=self.pivot_tree.xview)
        self.pivot_tree.configure(yscrollcommand=p_vsb.set, xscrollcommand=p_hsb.set)

        self.pivot_tree.grid(row=0, column=0, sticky="nsew")
        p_vsb.grid(row=0, column=1, sticky="ns")
        p_hsb.grid(row=1, column=0, sticky="ew")
        pivot_table_frame.grid_rowconfigure(0, weight=1)
        pivot_table_frame.grid_columnconfigure(0, weight=1)

    def apply_columns_config(self):
        is_enriched = (self.data_format == "CASE_ENRICHED")
        is_tin_bao = (self.data_format in ("TIN_BAO", "STANDARDIZED_TIN_BAO"))
        is_reconcile = bool(self.records and any(r.get("nguon_du_lieu") for r in self.records)) and not is_enriched
        is_case = (self.data_format in ("CASE", "CASE_REPORT", "CASE_ENRICHED", "STANDARDIZED_CASE") or (is_reconcile and not is_tin_bao))

        if is_enriched:
            cols_def = COLUMNS_CASE_27
        elif is_tin_bao and is_reconcile:
            cols_def = COLUMNS_TIN_BAO_RECONCILE_36
        elif is_tin_bao:
            cols_def = COLUMNS_TIN_BAO_35
        elif is_reconcile:
            cols_def = COLUMNS_CASE_27
        elif is_case:
            cols_def = COLUMNS_CASE_24
        else:
            cols_def = COLUMNS_12

        cols = [k for k, _, _ in cols_def]
        self.tree["columns"] = cols
        for key, title, width in cols_def:
            self.tree.heading(key, text=title)
            anchor = tk.CENTER if key in (
                "stt", "nguon_du_lieu", "tinh_thanh", "giai_doan", "ngay_sinh", "so_cccd", "so_qd_khoi_to",
                "ngay_qd_khoi_to", "ngay_tiep_nhan", "thoi_han_giai_quyet", "ngay_het_thoi_hieu",
                "so_ngay_con_lai", "muc_do_khan_cap", "canh_bao_thoi_hieu", "canh_bao_chi_tiet_thoi_hieu",
                "so_ngay_con_lai_thoi_hieu", "so_thu_ly", "ngay_thu_ly",
                "loai_thu_ly", "tong_bi_can", "loai_toi", "trang_thai_vu_an", "trang_thai",
                "ma_tin_bao", "loai_tin", "nguon_tin", "tinh_tiet_phuc_tap", "cccd_nguoi_bao", "sdt_nguoi_bao",
                "ngay_tiep_nhan_tin", "ngay_cqdt_nhan", "ngay_vks_nhan", "ngay_vks_thu_ly",
                "thoi_han_tu_ngay", "thoi_han_den_ngay", "ket_qua_giai_quyet", "ngay_tam_dinh_chi"
            ) else tk.W
            self.tree.column(key, width=width, anchor=anchor)

        if is_enriched:
            tab1_txt = "📋 Danh Sách Vụ Án Đã Liên Thông & Làm Giàu Bị Can"
            tab2_txt = "📊 Bảng Pivot Table Ma Trận (VKS Thụ Lý × Giai Đoạn)"
        elif is_tin_bao and is_reconcile:
            tab1_txt = "📋 Danh Sách 39 Cột Hợp Nhất & Đối Chiếu Tin Báo Đa Nguồn QLAHS"
            tab2_txt = "📊 Bảng Pivot Table Ma Trận (VKS Thụ Lý × Kết Quả Giải Quyết)"
        elif is_tin_bao:
            tab1_txt = "📋 Danh Sách 38 Cột Chuẩn Hóa (Nguồn Tin Báo & Tố Giác Tội Phạm)"
            tab2_txt = "📊 Bảng Pivot Table Ma Trận (VKS Thụ Lý × Kết Quả Giải Quyết)"
        elif is_reconcile:
            tab1_txt = "📋 Danh Sách 27 Cột Hợp Nhất & Đối Chiếu Đa Nguồn QLAHS"
            tab2_txt = "📊 Bảng Pivot Table Ma Trận (VKS Thụ Lý × Giai Đoạn)"
        elif is_case:
            tab1_txt = "📋 Danh Sách 24 Cột Chuẩn Hóa (Sổ Vụ Án & Thụ Lý)"
            tab2_txt = "📊 Bảng Pivot Table Ma Trận (VKS Thụ Lý × Giai Đoạn)"
        else:
            tab1_txt = "📋 Danh Sách 12 Cột Chuẩn Hóa (Bao gồm Tỉnh/Thành)"
            tab2_txt = "📊 Bảng Pivot Table Ma Trận (Tỉnh/Thành × Giai Đoạn)"

        self.notebook.tab(self.tab_table, text=tab1_txt)
        self.notebook.tab(self.tab_pivot, text=tab2_txt)

        if hasattr(self, "cbo_pivot_mode"):
            if is_tin_bao:
                p_modes = [
                    "📊 1. MA TRẬN KẾT QUẢ: Đơn Vị VKS × Kết Quả Giải Quyết Tin Báo",
                    "🚨 2. MA TRẬN RỦI RO: Đơn Vị VKS × Cảnh Báo Thời Hiệu & Quá Hạn (Đ.230)",
                    "👑 3. BÁO CÁO TOÀN DIỆN 360°: Tổng Quan Điều Hành Tin Báo Gửi Lãnh Đạo",
                    "👤 4. PHÂN CÔNG NGHIỆP VỤ: Khối Lượng Thụ Lý & Tình Hình Giải Quyết Theo KSV"
                ]
            else:
                p_modes = [
                    "🏛️ 1. MA TRẬN TIẾN ĐỘ: Đơn Vị VKS × Giai Đoạn Tố Tụng (Điều Tra, Truy Tố, Xét Xử)",
                    "🚨 2. MA TRẬN RỦI RO: Đơn Vị VKS × Cảnh Báo Thời Hiệu TNHS & Quá Hạn (Đ.230)",
                    "👑 3. BÁO CÁO TOÀN DIỆN 360°: Tổng Quan Điều Hành Lãnh Đạo Viện",
                    "👤 4. PHÂN CÔNG NGHIỆP VỤ: Kiểm Sát Viên × Trạng Thái Vụ Án",
                    "⚖️ 5. PHÂN LOẠI TỘI PHẠM: Đơn Vị VKS × Tính Chất Tội (Ít NT, NT, RNT, ĐBNT)"
                ]
            curr_m = self.cbo_pivot_mode.get()
            self.cbo_pivot_mode["values"] = p_modes
            if curr_m not in p_modes:
                self.cbo_pivot_mode.current(0)

    def get_filtered_records(self):
        q = self.search_var.get().strip().lower()
        prov = self.prov_filter_var.get()
        stg = self.stage_filter_var.get()
        sts = self.status_filter_var.get()

        is_tin_bao = (self.data_format in ("TIN_BAO", "STANDARDIZED_TIN_BAO"))
        is_reconcile = bool(self.records and any(r.get("nguon_du_lieu") for r in self.records))
        is_case = (self.data_format in ("CASE", "CASE_REPORT", "CASE_ENRICHED", "STANDARDIZED_CASE") or is_reconcile)

        out = []
        for r in self.records:
            if is_tin_bao:
                if prov != "Tất cả" and r.get("vks_thu_ly", "") != prov:
                    continue
                if stg != "Tất cả" and r.get("ket_qua_giai_quyet", "") != stg:
                    continue
                if sts != "Tất cả":
                    urg_field = r.get("muc_do_khan_cap", "")
                    if sts == "Quá hạn" and "🔴" not in urg_field:
                        continue
                    elif sts == "Khẩn cấp" and "🟠" not in urg_field:
                        continue
                    elif sts == "Cảnh báo" and "🟡" not in urg_field:
                        continue
                    elif sts == "Trong thời hạn" and "🟢" not in urg_field:
                        continue
                    elif sts == "Tạm đình chỉ" and "⏸" not in urg_field:
                        continue
                    elif sts == "Đã giải quyết" and "✔" not in urg_field:
                        continue
                if q:
                    comb = " ".join(str(v or "") for v in r.values()).lower()
                    if q not in comb:
                        continue
                out.append(r)
            elif is_case:
                if prov != "Tất cả" and r.get("tinh_thanh", "") != prov:
                    continue
                if stg != "Tất cả" and r.get("giai_doan", "").lower() != stg.lower():
                    continue
                if sts != "Tất cả":
                    r_st = r.get("trang_thai_vu_an", "").lower()
                    if sts.lower() not in r_st:
                        continue
                if q:
                    line = " ".join(str(v) for v in r.values()).lower()
                    if q not in line:
                        continue
                out.append(r)
            else:
                if prov != "Tất cả" and r.get("tinh_thanh", "") != prov:
                    continue
                if stg != "Tất cả" and r.get("giai_doan", "").lower() != stg.lower():
                    continue
                if sts != "Tất cả":
                    r_st = r.get("trang_thai", "").lower()
                    if sts.lower() not in r_st:
                        continue
                if q:
                    line = " ".join(str(v) for v in r.values()).lower()
                    if q not in line:
                        continue
                out.append(r)
        return out

    def update_filter_options(self):
        is_tin_bao = (self.data_format in ("TIN_BAO", "STANDARDIZED_TIN_BAO"))
        is_case = (self.data_format in ("CASE", "CASE_REPORT", "CASE_ENRICHED", "STANDARDIZED_CASE"))

        if is_tin_bao:
            units = sorted(list(set(r.get("vks_thu_ly", "").strip() for r in self.records if r.get("vks_thu_ly", "").strip())))
            self.cbo_prov["values"] = ["Tất cả"] + units
            if self.prov_filter_var.get() not in self.cbo_prov["values"]:
                self.prov_filter_var.set("Tất cả")

            kqs = sorted(list(set(r.get("ket_qua_giai_quyet", "").strip() for r in self.records if r.get("ket_qua_giai_quyet", "").strip())))
            self.cbo_stage["values"] = ["Tất cả"] + kqs
            if self.stage_filter_var.get() not in self.cbo_stage["values"]:
                self.stage_filter_var.set("Tất cả")

            self.cbo_status["values"] = [
                "Tất cả",
                "Quá hạn",
                "Khẩn cấp",
                "Cảnh báo",
                "Trong thời hạn",
                "Tạm đình chỉ",
                "Đã giải quyết"
            ]
            if self.status_filter_var.get() not in self.cbo_status["values"]:
                self.status_filter_var.set("Tất cả")
        elif is_case:
            unique_provs = sorted(list(set(r.get("tinh_thanh", "").strip() for r in self.records if r.get("tinh_thanh", "").strip())))
            self.cbo_prov["values"] = ["Tất cả"] + unique_provs
            if self.prov_filter_var.get() not in self.cbo_prov["values"]:
                self.prov_filter_var.set("Tất cả")

            stages = sorted(list(set(r.get("giai_doan", "").strip() for r in self.records if r.get("giai_doan", "").strip())))
            self.cbo_stage["values"] = ["Tất cả"] + stages
            if self.stage_filter_var.get() not in self.cbo_stage["values"]:
                self.stage_filter_var.set("Tất cả")

            statuses = sorted(list(set(r.get("trang_thai_vu_an", "").strip() for r in self.records if r.get("trang_thai_vu_an", "").strip())))
            self.cbo_status["values"] = ["Tất cả"] + statuses
            if self.status_filter_var.get() not in self.cbo_status["values"]:
                self.status_filter_var.set("Tất cả")
        else:
            unique_provs = sorted(list(set(r.get("tinh_thanh", "").strip() for r in self.records if r.get("tinh_thanh", "").strip())))
            self.cbo_prov["values"] = ["Tất cả"] + unique_provs
            if self.prov_filter_var.get() not in self.cbo_prov["values"]:
                self.prov_filter_var.set("Tất cả")

            self.cbo_stage["values"] = ["Tất cả", "Điều tra", "Truy tố", "Sơ thẩm", "Phúc thẩm"]
            if self.stage_filter_var.get() not in self.cbo_stage["values"]:
                self.stage_filter_var.set("Tất cả")

            self.cbo_status["values"] = ["Tất cả", "Đình nã", "Truy nã", "Tạm đình chỉ"]
            if self.status_filter_var.get() not in self.cbo_status["values"]:
                self.status_filter_var.set("Tất cả")

    def refresh_views(self):
        self.apply_columns_config()
        filtered = self.get_filtered_records()
        self.update_filter_options()

        is_enriched = (self.data_format == "CASE_ENRICHED")
        is_tin_bao = (self.data_format in ("TIN_BAO", "STANDARDIZED_TIN_BAO"))
        is_reconcile = bool(self.records and any(r.get("nguon_du_lieu") for r in self.records)) and not is_enriched
        is_case = (self.data_format in ("CASE", "CASE_REPORT", "CASE_ENRICHED", "STANDARDIZED_CASE") or (is_reconcile and not is_tin_bao))

        if is_enriched:
            cols_def = COLUMNS_CASE_27
        elif is_tin_bao and is_reconcile:
            cols_def = COLUMNS_TIN_BAO_RECONCILE_36
        elif is_tin_bao:
            cols_def = COLUMNS_TIN_BAO_35
        elif is_reconcile:
            cols_def = COLUMNS_CASE_27
        elif is_case:
            cols_def = COLUMNS_CASE_24
        else:
            cols_def = COLUMNS_12

        # 1. Cập nhật bảng dữ liệu với màu xen kẽ
        for item in self.tree.get_children():
            self.tree.delete(item)

        for idx, r in enumerate(filtered):
            vals = [r.get(k, "") for k, _, _ in cols_def]
            row_tag = "even" if idx % 2 == 0 else "odd"
            tags = [row_tag]
            if is_tin_bao:
                urg_str = r.get("muc_do_khan_cap", "")
                if "🔴" in urg_str:
                    tags.append("tag_khan_cap")
                elif "🟠" in urg_str:
                    tags.append("tag_rat_gap")
                elif "🟡" in urg_str:
                    tags.append("tag_sap_het")
                elif "⏸" in urg_str:
                    tags.append("dinh_na")
            elif is_case:
                urg_str = r.get("muc_do_khan_cap", "")
                if "🔴" in urg_str:
                    tags.append("tag_khan_cap")
                elif "🟠" in urg_str:
                    tags.append("tag_rat_gap")
                elif "🟡" in urg_str:
                    tags.append("tag_sap_het")
                elif r.get("canh_bao_thoi_hieu"):
                    tags.append("sap_het_thoi_hieu")
                if "tạm đình chỉ" in str(r.get("trang_thai_vu_an", "")).lower():
                    tags.append("dinh_na")
            else:
                tt = str(r.get("trang_thai", "")).lower()
                if "đình nã" in tt:
                    tags.append("dinh_na")
                elif "truy nã" in tt:
                    tags.append("truy_na")

            self.tree.insert("", tk.END, values=vals, tags=tuple(tags))

        if is_tin_bao:
            unit_label = "nguồn tin báo & tố giác tội phạm"
        elif is_case:
            unit_label = "vụ án"
        else:
            unit_label = "hồ sơ bị can"

        self.lbl_status.config(
            text=f"Đã tự động nhận diện và chuẩn hóa {len(self.records):,} {unit_label} (Đang hiển thị {len(filtered):,} bản ghi theo bộ lọc)."
        )

        # 2. Cập nhật thẻ KPI Điều Hành Lãnh Đạo
        if is_tin_bao:
            total_tb = len(self.records)
            expired_tb = sum(1 for r in self.records if "⚠️" in str(r.get("muc_do_khan_cap") or "") or "hết thời hiệu" in str(r.get("muc_do_khan_cap") or "").lower())
            urgent_tb = sum(1 for r in self.records if "🔴" in str(r.get("muc_do_khan_cap") or "") or "quá hạn" in str(r.get("muc_do_khan_cap") or "").lower())
            suspended_tb = sum(1 for r in self.records if "tạm đình chỉ" in str(r.get("ket_qua_giai_quyet") or "").lower() or "⏸" in str(r.get("muc_do_khan_cap") or ""))
            solved_tb = sum(1 for r in self.records if any(k in str(r.get("ket_qua_giai_quyet") or "").lower() for k in ("khởi tố", "không khởi tố", "đã nhập", "chuyển nơi khác")))

            self.kpi_titles["total"].config(text="TỔNG SỐ TIN BÁO")
            self.kpi_cards["total"].config(text=f"{total_tb:,} tin báo")
            self.kpi_titles["cases"].config(text="⚠️ ĐÃ HẾT THỜI HIỆU (Đ.230)")
            self.kpi_cards["cases"].config(text=f"{expired_tb:,} tin báo")
            self.kpi_titles["provinces"].config(text="🔴 KHẨN CẤP (≤ 7 NGÀY)")
            self.kpi_cards["provinces"].config(text=f"{urgent_tb:,} tin báo")
            self.kpi_titles["stages"].config(text="⏸️ TẠM ĐÌNH CHỈ GIẢI QUYẾT")
            self.kpi_cards["stages"].config(text=f"{suspended_tb:,} tin báo")
            self.kpi_titles["status"].config(text="✔ ĐÃ HOÀN THÀNH TỐ TỤNG")
            pct_s = (solved_tb / total_tb * 100) if total_tb else 0
            self.kpi_cards["status"].config(text=f"{solved_tb:,} tin ({pct_s:.1f}%)")
        elif is_case:
            total_cases = len(self.records)
            expired_statute = sum(1 for r in self.records if "⚠️" in str(r.get("muc_do_khan_cap") or "") or "hết thời hiệu" in str(r.get("muc_do_khan_cap") or "").lower())
            urgent_cnt = sum(1 for r in self.records if "🔴" in str(r.get("muc_do_khan_cap") or "") or "🟠" in str(r.get("muc_do_khan_cap") or ""))
            suspended_cnt = sum(1 for r in self.records if "tạm đình chỉ" in str(r.get("trang_thai_vu_an") or "").lower() or "⏸" in str(r.get("muc_do_khan_cap") or ""))
            dt_cnt = sum(1 for r in self.records if str(r.get("giai_doan") or "").strip() == "Điều tra")
            tt_cnt = sum(1 for r in self.records if str(r.get("giai_doan") or "").strip() == "Truy tố")
            st_cnt = sum(1 for r in self.records if str(r.get("giai_doan") or "").strip() == "Sơ thẩm")

            self.kpi_titles["total"].config(text="TỔNG SỐ ÁN QUẢN LÝ")
            self.kpi_cards["total"].config(text=f"{total_cases:,} vụ án")
            self.kpi_titles["cases"].config(text="⚠️ ĐÃ HẾT THỜI HIỆU (Đ.230)")
            self.kpi_cards["cases"].config(text=f"{expired_statute:,} vụ án")
            self.kpi_titles["provinces"].config(text="🔴 KHẨN CẤP (≤ 15 NGÀY)")
            self.kpi_cards["provinces"].config(text=f"{urgent_cnt:,} vụ án")
            self.kpi_titles["stages"].config(text="⏸️ TẠM ĐÌNH CHỈ / ĐÌNH CHỈ")
            self.kpi_cards["stages"].config(text=f"{suspended_cnt:,} vụ án")
            self.kpi_titles["status"].config(text="⚖️ PHÂN BỔ GIAI ĐOẠN")
            self.kpi_cards["status"].config(text=f"ĐT: {dt_cnt:,} | TT: {tt_cnt:,} | ST: {st_cnt:,}")
        else:
            total_defs = len(self.records)
            unique_cases = len(set(r.get("ma_vu_an", "").strip() for r in self.records if r.get("ma_vu_an", "").strip()))
            unique_provs = len(set(r.get("tinh_thanh", "").strip() for r in self.records if r.get("tinh_thanh", "").strip()))
            stg_counts = {}
            sts_counts = {}
            for r in self.records:
                g = r.get("giai_doan", "").strip() or "Chưa rõ"
                stg_counts[g] = stg_counts.get(g, 0) + 1
                s = r.get("trang_thai", "").strip() or "Khác"
                sts_counts[s] = sts_counts.get(s, 0) + 1

            self.kpi_titles["total"].config(text="TỔNG SỐ BỊ CAN")
            self.kpi_cards["total"].config(text=f"{total_defs:,} bị can")
            self.kpi_titles["cases"].config(text="SỐ VỤ ÁN LIÊN QUAN")
            self.kpi_cards["cases"].config(text=f"{unique_cases:,} vụ án")
            self.kpi_titles["provinces"].config(text="ĐỊA BÀN / TỈNH THÀNH")
            self.kpi_cards["provinces"].config(text=f"{unique_provs} tỉnh/thành")
            self.kpi_titles["stages"].config(text="GIAI ĐOẠN TỐ TỤNG")
            self.kpi_cards["stages"].config(text=" | ".join(f"{k}: {v}" for k, v in stg_counts.items()) or "Chưa có")
            self.kpi_titles["status"].config(text="TRẠNG THÁI HỒ SƠ")
            self.kpi_cards["status"].config(text=" | ".join(f"{k}: {v}" for k, v in sts_counts.items()) or "Chưa có")

        # 3. Cập nhật bảng Pivot Table
        for item in self.pivot_tree.get_children():
            self.pivot_tree.delete(item)

        filter_defs = len(filtered)
        mode = self.cbo_pivot_mode.get() if hasattr(self, "cbo_pivot_mode") else ""
        mode_u = mode.upper()

        if is_tin_bao:
            if "RỦI RO" in mode_u or "THỜI HIỆU" in mode_u:
                # Chế độ 2: VKS Thụ Lý × Cảnh Báo Thời Hiệu & Quá Hạn (Đ.230 BLTTHS)
                unit_stats, _, sorted_units, tot_all = build_executive_tin_bao_pivot_summary(filtered)
                cols = ["stt", "vks_thu_ly", "so_vu", "qua_han", "khan_cap", "rat_gap", "sap_het", "trong_han", "tam_dinh_chi", "ty_le"]
                self.pivot_tree["columns"] = cols
                headers_urg = [
                    ("stt", "STT", 50, tk.CENTER),
                    ("vks_thu_ly", "Đơn vị VKS thụ lý", 260, tk.W),
                    ("so_vu", "Tổng số tin", 90, tk.CENTER),
                    ("qua_han", "Hết thời hiệu ⚠️ (Đ.230)", 160, tk.CENTER),
                    ("khan_cap", "🔴 Khẩn cấp (≤7 ngày)", 140, tk.CENTER),
                    ("rat_gap", "🟠 Rất gấp (8-15 ngày)", 140, tk.CENTER),
                    ("sap_het", "🟡 Sắp hết hạn (16-30 ngày)", 150, tk.CENTER),
                    ("trong_han", "🟢 Trong hạn (>30 ngày)", 140, tk.CENTER),
                    ("tam_dinh_chi", "⏸️ Tạm đình chỉ", 110, tk.CENTER),
                    ("ty_le", "Tỷ lệ %", 85, tk.CENTER),
                ]
                for cid, htxt, cw, c_align in headers_urg:
                    self.pivot_tree.heading(cid, text=htxt)
                    self.pivot_tree.column(cid, width=cw, anchor=c_align)

                for idx, u in enumerate(sorted_units, 1):
                    st = unit_stats[u]
                    pct = (st["total_tb"] / tot_all * 100) if tot_all > 0 else 0
                    row_vals = [
                        idx, u, st["total_tb"],
                        st["qua_han"] or "",
                        st["khan_cap"] or "",
                        st["rat_gap"] or "",
                        st["sap_het"] or "",
                        st["binh_thuong"] or "",
                        st["tam_dinh_chi"] or "",
                        f"{pct:.1f}%",
                    ]
                    row_tag = "p_even" if idx % 2 == 0 else "p_odd"
                    self.pivot_tree.insert("", tk.END, values=row_vals, tags=(row_tag,))

                if sorted_units:
                    tot_row = [
                        "∑", "TỔNG CỘNG TOÀN VIỆN", tot_all,
                        sum(unit_stats[u]["qua_han"] for u in sorted_units),
                        sum(unit_stats[u]["khan_cap"] for u in sorted_units),
                        sum(unit_stats[u]["rat_gap"] for u in sorted_units),
                        sum(unit_stats[u]["sap_het"] for u in sorted_units),
                        sum(unit_stats[u]["binh_thuong"] for u in sorted_units),
                        sum(unit_stats[u]["tam_dinh_chi"] for u in sorted_units),
                        "100.0%",
                    ]
                    self.pivot_tree.insert("", tk.END, values=tot_row, tags=("p_total",))

            elif "TOÀN DIỆN" in mode_u or "360" in mode:
                # Chế độ 3: Báo Cáo Điều Hành 360° Gửi Lãnh Đạo
                unit_stats, active_outcomes, sorted_units, tot_all = build_executive_tin_bao_pivot_summary(filtered)
                cols = ["stt", "vks_thu_ly"] + [f"oc_{i}" for i in range(len(active_outcomes))] + ["qua_han", "khan_cap", "so_vu", "ty_le"]
                self.pivot_tree["columns"] = cols
                self.pivot_tree.heading("stt", text="STT")
                self.pivot_tree.column("stt", width=50, anchor=tk.CENTER)
                self.pivot_tree.heading("vks_thu_ly", text="Đơn vị VKS thụ lý")
                self.pivot_tree.column("vks_thu_ly", width=250, anchor=tk.W)
                for i, oc in enumerate(active_outcomes):
                    cid = f"oc_{i}"
                    self.pivot_tree.heading(cid, text=oc)
                    self.pivot_tree.column(cid, width=110, anchor=tk.CENTER)
                self.pivot_tree.heading("qua_han", text="Hết thời hiệu ⚠️")
                self.pivot_tree.column("qua_han", width=110, anchor=tk.CENTER)
                self.pivot_tree.heading("khan_cap", text="Khẩn cấp 🔴")
                self.pivot_tree.column("khan_cap", width=100, anchor=tk.CENTER)
                self.pivot_tree.heading("so_vu", text="Tổng số tin")
                self.pivot_tree.column("so_vu", width=90, anchor=tk.CENTER)
                self.pivot_tree.heading("ty_le", text="Tỷ lệ %")
                self.pivot_tree.column("ty_le", width=80, anchor=tk.CENTER)

                for idx, u in enumerate(sorted_units, 1):
                    st = unit_stats[u]
                    pct = (st["total_tb"] / tot_all * 100) if tot_all > 0 else 0
                    row_vals = [idx, u] + [st["outcomes"].get(k, 0) or "" for k in active_outcomes] + [st["qua_han"] or "", st["khan_cap"] or "", st["total_tb"], f"{pct:.1f}%"]
                    row_tag = "p_even" if idx % 2 == 0 else "p_odd"
                    self.pivot_tree.insert("", tk.END, values=row_vals, tags=(row_tag,))

                if sorted_units:
                    tot_row = ["∑", "TỔNG CỘNG TOÀN VIỆN"] + [sum(unit_stats[u]["outcomes"].get(k, 0) for u in sorted_units) for k in active_outcomes] + [
                        sum(unit_stats[u]["qua_han"] for u in sorted_units),
                        sum(unit_stats[u]["khan_cap"] for u in sorted_units),
                        tot_all, "100.0%"
                    ]
                    self.pivot_tree.insert("", tk.END, values=tot_row, tags=("p_total",))

            elif "PHÂN CÔNG" in mode_u or "KIỂM SÁT VIÊN" in mode_u or "KSV" in mode_u:
                # Chế độ 4: Phân công nghiệp vụ theo KSV
                officer_stats, sorted_officers, tot_all = build_officer_assignment_pivot_summary(filtered, is_tin_bao=True)
                cols = ["stt", "officer", "so_vu", "qua_han", "khan_cap", "tam_dinh_chi", "hoan_thanh", "ty_le"]
                self.pivot_tree["columns"] = cols
                headers_ksv = [
                    ("stt", "STT", 50, tk.CENTER),
                    ("officer", "Kiểm Sát Viên thụ lý", 240, tk.W),
                    ("so_vu", "Tổng số tin thụ lý", 120, tk.CENTER),
                    ("qua_han", "Quá hạn / Hết TH ⚠️", 130, tk.CENTER),
                    ("khan_cap", "Khẩn cấp 🔴", 100, tk.CENTER),
                    ("tam_dinh_chi", "Tạm đình chỉ ⏸️", 110, tk.CENTER),
                    ("hoan_thanh", "Đã giải quyết ✔", 110, tk.CENTER),
                    ("ty_le", "Tỷ lệ %", 85, tk.CENTER),
                ]
                for cid, htxt, cw, c_align in headers_ksv:
                    self.pivot_tree.heading(cid, text=htxt)
                    self.pivot_tree.column(cid, width=cw, anchor=c_align)

                for idx, off in enumerate(sorted_officers, 1):
                    st = officer_stats[off]
                    pct = (st["total"] / tot_all * 100) if tot_all > 0 else 0
                    row_vals = [idx, off, st["total"], st["qua_han"] or "", st["khan_cap"] or "", st["tam_dinh_chi"] or "", st["hoan_thanh"] or "", f"{pct:.1f}%"]
                    row_tag = "p_even" if idx % 2 == 0 else "p_odd"
                    self.pivot_tree.insert("", tk.END, values=row_vals, tags=(row_tag,))

                if sorted_officers:
                    tot_row = [
                        "∑", "TỔNG CỘNG TOÀN VIỆN", tot_all,
                        sum(officer_stats[o]["qua_han"] for o in sorted_officers),
                        sum(officer_stats[o]["khan_cap"] for o in sorted_officers),
                        sum(officer_stats[o]["tam_dinh_chi"] for o in sorted_officers),
                        sum(officer_stats[o]["hoan_thanh"] for o in sorted_officers),
                        "100.0%"
                    ]
                    self.pivot_tree.insert("", tk.END, values=tot_row, tags=("p_total",))

            else:
                # Chế độ 1: VKS Thụ Lý × Kết Quả Giải Quyết
                matrix, sorted_results, sorted_units = build_pivot_tin_bao_summary(filtered)
                cols = ["stt", "vks_thu_ly", "so_vu"] + [f"kq_{i}" for i in range(len(sorted_results))] + ["ty_le"]
                self.pivot_tree["columns"] = cols
                self.pivot_tree.heading("stt", text="STT")
                self.pivot_tree.column("stt", width=50, anchor=tk.CENTER)
                self.pivot_tree.heading("vks_thu_ly", text="Đơn vị VKS thụ lý")
                self.pivot_tree.column("vks_thu_ly", width=250, anchor=tk.W)
                self.pivot_tree.heading("so_vu", text="Tổng số tin")
                self.pivot_tree.column("so_vu", width=90, anchor=tk.CENTER)
                for i, kq in enumerate(sorted_results):
                    cid = f"kq_{i}"
                    self.pivot_tree.heading(cid, text=kq)
                    self.pivot_tree.column(cid, width=120, anchor=tk.CENTER)
                self.pivot_tree.heading("ty_le", text="Tỷ lệ %")
                self.pivot_tree.column("ty_le", width=85, anchor=tk.CENTER)

                for idx, u in enumerate(sorted_units, 1):
                    data = matrix[u]
                    tot = data["_total"]
                    pct = (tot / filter_defs * 100) if filter_defs > 0 else 0
                    row_vals = [idx, u, tot] + [data.get(k, 0) for k in sorted_results] + [f"{pct:.1f}%"]
                    row_tag = "p_even" if idx % 2 == 0 else "p_odd"
                    self.pivot_tree.insert("", tk.END, values=row_vals, tags=(row_tag,))

                if sorted_units:
                    tot_row = ["∑", "TỔNG CỘNG CHUNG", filter_defs] + [sum(matrix[u].get(k, 0) for u in sorted_units) for k in sorted_results] + ["100.0%"]
                    self.pivot_tree.insert("", tk.END, values=tot_row, tags=("p_total",))
        elif is_case:
            mode = self.cbo_pivot_mode.get() if hasattr(self, "cbo_pivot_mode") else ""
            mode_u = mode.upper()

            if "TOÀN DIỆN" in mode_u or "360" in mode or "ĐA CHIỀU" in mode_u:
                # Chế độ 3: Báo Cáo Tổng Quan Đa Chiều 360° Gửi Lãnh Đạo
                unit_stats, active_stages, sorted_units, tot_all = build_executive_master_pivot_summary(filtered)
                cols = [
                    "stt", "vks_thu_ly", "dieu_tra", "truy_to", "so_tham",
                    "qua_han", "khan_cap", "rat_gap", "sap_het", "trong_han", "tam_dinh_chi",
                    "it_nt", "nt", "rat_nt", "dbnt",
                    "tong_so_vu", "tong_bi_can", "ty_le"
                ]
                self.pivot_tree["columns"] = cols
                headers_info = [
                    ("stt", "STT", 45, tk.CENTER),
                    ("vks_thu_ly", "Đơn vị VKS thụ lý", 240, tk.W),
                    ("dieu_tra", "Điều tra", 75, tk.CENTER),
                    ("truy_to", "Truy tố", 75, tk.CENTER),
                    ("so_tham", "Sơ thẩm", 75, tk.CENTER),
                    ("qua_han", "Quá hạn ⚠️", 80, tk.CENTER),
                    ("khan_cap", "Khẩn cấp 🔴", 85, tk.CENTER),
                    ("rat_gap", "Rất gấp 🟠", 80, tk.CENTER),
                    ("sap_het", "Sắp hết 🟡", 80, tk.CENTER),
                    ("trong_han", "Trong hạn 🟢", 80, tk.CENTER),
                    ("tam_dinh_chi", "Tạm ĐC ⏸️", 85, tk.CENTER),
                    ("it_nt", "Ít NT", 70, tk.CENTER),
                    ("nt", "Nghiêm trọng", 85, tk.CENTER),
                    ("rat_nt", "Rất NT", 70, tk.CENTER),
                    ("dbnt", "Đặc biệt NT", 85, tk.CENTER),
                    ("tong_so_vu", "Tổng số vụ", 85, tk.CENTER),
                    ("tong_bi_can", "Tổng bị can", 85, tk.CENTER),
                    ("ty_le", "Tỷ lệ %", 75, tk.CENTER),
                ]
                for cid, htxt, cw, c_align in headers_info:
                    self.pivot_tree.heading(cid, text=htxt)
                    self.pivot_tree.column(cid, width=cw, anchor=c_align)

                for idx, u in enumerate(sorted_units, 1):
                    st = unit_stats[u]
                    pct = (st["total_cases"] / tot_all * 100) if tot_all > 0 else 0
                    row_vals = [
                        idx, u,
                        st["stages"].get("Điều tra", 0) or "",
                        st["stages"].get("Truy tố", 0) or "",
                        st["stages"].get("Sơ thẩm", 0) or "",
                        st["qua_han"] or "",
                        st["khan_cap"] or "",
                        st["rat_gap"] or "",
                        st["sap_het"] or "",
                        st["binh_thuong"] or "",
                        st["tam_dinh_chi"] or "",
                        st["toi_it"] or "",
                        st["toi_nt"] or "",
                        st["toi_rnt"] or "",
                        st["toi_dbnt"] or "",
                        st["total_cases"],
                        st["total_defendants"],
                        f"{pct:.1f}%",
                    ]
                    row_tag = "p_even" if idx % 2 == 0 else "p_odd"
                    self.pivot_tree.insert("", tk.END, values=row_vals, tags=(row_tag,))

                if sorted_units:
                    tot_row = [
                        "∑", "Grand Total (TỔNG CỘNG TOÀN VIỆN)",
                        sum(unit_stats[u]["stages"].get("Điều tra", 0) for u in sorted_units),
                        sum(unit_stats[u]["stages"].get("Truy tố", 0) for u in sorted_units),
                        sum(unit_stats[u]["stages"].get("Sơ thẩm", 0) for u in sorted_units),
                        sum(unit_stats[u]["qua_han"] for u in sorted_units),
                        sum(unit_stats[u]["khan_cap"] for u in sorted_units),
                        sum(unit_stats[u]["rat_gap"] for u in sorted_units),
                        sum(unit_stats[u]["sap_het"] for u in sorted_units),
                        sum(unit_stats[u]["binh_thuong"] for u in sorted_units),
                        sum(unit_stats[u]["tam_dinh_chi"] for u in sorted_units),
                        sum(unit_stats[u]["toi_it"] for u in sorted_units),
                        sum(unit_stats[u]["toi_nt"] for u in sorted_units),
                        sum(unit_stats[u]["toi_rnt"] for u in sorted_units),
                        sum(unit_stats[u]["toi_dbnt"] for u in sorted_units),
                        tot_all,
                        sum(unit_stats[u]["total_defendants"] for u in sorted_units),
                        "100.0%",
                    ]
                    self.pivot_tree.insert("", tk.END, values=tot_row, tags=("p_total",))

            elif "RỦI RO" in mode_u or "THỜI HIỆU" in mode_u:
                # Chế độ 2: VKS Thụ Lý × Cảnh Báo Thời Hiệu & Khẩn Cấp (Đ.230 BLTTHS)
                unit_stats, _, sorted_units, tot_all = build_executive_master_pivot_summary(filtered)
                cols = ["stt", "vks_thu_ly", "so_vu", "qua_han", "khan_cap", "rat_gap", "sap_het", "trong_han", "tam_dinh_chi", "ty_le"]
                self.pivot_tree["columns"] = cols
                headers_urg = [
                    ("stt", "STT", 50, tk.CENTER),
                    ("vks_thu_ly", "Đơn vị VKS thụ lý", 260, tk.W),
                    ("so_vu", "Tổng số vụ", 90, tk.CENTER),
                    ("qua_han", "Hết thời hiệu ⚠️ (Đ.230)", 160, tk.CENTER),
                    ("khan_cap", "🔴 Khẩn cấp (≤7 ngày)", 140, tk.CENTER),
                    ("rat_gap", "🟠 Rất gấp (8-15 ngày)", 140, tk.CENTER),
                    ("sap_het", "🟡 Sắp hết hạn (16-30 ngày)", 150, tk.CENTER),
                    ("trong_han", "🟢 Trong hạn (>30 ngày)", 140, tk.CENTER),
                    ("tam_dinh_chi", "⏸️ Tạm đình chỉ", 110, tk.CENTER),
                    ("ty_le", "Tỷ lệ %", 85, tk.CENTER),
                ]
                for cid, htxt, cw, c_align in headers_urg:
                    self.pivot_tree.heading(cid, text=htxt)
                    self.pivot_tree.column(cid, width=cw, anchor=c_align)

                for idx, u in enumerate(sorted_units, 1):
                    st = unit_stats[u]
                    pct = (st["total_cases"] / tot_all * 100) if tot_all > 0 else 0
                    row_vals = [
                        idx, u, st["total_cases"],
                        st["qua_han"] or "",
                        st["khan_cap"] or "",
                        st["rat_gap"] or "",
                        st["sap_het"] or "",
                        st["binh_thuong"] or "",
                        st["tam_dinh_chi"] or "",
                        f"{pct:.1f}%",
                    ]
                    row_tag = "p_even" if idx % 2 == 0 else "p_odd"
                    self.pivot_tree.insert("", tk.END, values=row_vals, tags=(row_tag,))

                if sorted_units:
                    tot_row = [
                        "∑", "TỔNG CỘNG TOÀN VIỆN", tot_all,
                        sum(unit_stats[u]["qua_han"] for u in sorted_units),
                        sum(unit_stats[u]["khan_cap"] for u in sorted_units),
                        sum(unit_stats[u]["rat_gap"] for u in sorted_units),
                        sum(unit_stats[u]["sap_het"] for u in sorted_units),
                        sum(unit_stats[u]["binh_thuong"] for u in sorted_units),
                        sum(unit_stats[u]["tam_dinh_chi"] for u in sorted_units),
                        "100.0%",
                    ]
                    self.pivot_tree.insert("", tk.END, values=tot_row, tags=("p_total",))

            elif "PHÂN CÔNG" in mode_u or "KIỂM SÁT VIÊN" in mode_u or "KSV" in mode_u:
                # Chế độ 4: Phân công nghiệp vụ theo KSV (Vụ án)
                officer_stats, sorted_officers, tot_all = build_officer_assignment_pivot_summary(filtered, is_tin_bao=False)
                cols = ["stt", "officer", "so_vu", "qua_han", "khan_cap", "tam_dinh_chi", "hoan_thanh", "ty_le"]
                self.pivot_tree["columns"] = cols
                headers_ksv = [
                    ("stt", "STT", 50, tk.CENTER),
                    ("officer", "Kiểm Sát Viên giải quyết", 240, tk.W),
                    ("so_vu", "Tổng số vụ thụ lý", 120, tk.CENTER),
                    ("qua_han", "Hết thời hiệu / Quá hạn ⚠️", 150, tk.CENTER),
                    ("khan_cap", "Khẩn cấp 🔴 (≤15 ngày)", 130, tk.CENTER),
                    ("tam_dinh_chi", "Tạm đình chỉ ⏸️", 110, tk.CENTER),
                    ("hoan_thanh", "Đã giải quyết / Xét xử ✔", 150, tk.CENTER),
                    ("ty_le", "Tỷ lệ %", 85, tk.CENTER),
                ]
                for cid, htxt, cw, c_align in headers_ksv:
                    self.pivot_tree.heading(cid, text=htxt)
                    self.pivot_tree.column(cid, width=cw, anchor=c_align)

                for idx, off in enumerate(sorted_officers, 1):
                    st = officer_stats[off]
                    pct = (st["total"] / tot_all * 100) if tot_all > 0 else 0
                    row_vals = [idx, off, st["total"], st["qua_han"] or "", st["khan_cap"] or "", st["tam_dinh_chi"] or "", st["hoan_thanh"] or "", f"{pct:.1f}%"]
                    row_tag = "p_even" if idx % 2 == 0 else "p_odd"
                    self.pivot_tree.insert("", tk.END, values=row_vals, tags=(row_tag,))

                if sorted_officers:
                    tot_row = [
                        "∑", "TỔNG CỘNG TOÀN VIỆN", tot_all,
                        sum(officer_stats[o]["qua_han"] for o in sorted_officers),
                        sum(officer_stats[o]["khan_cap"] for o in sorted_officers),
                        sum(officer_stats[o]["tam_dinh_chi"] for o in sorted_officers),
                        sum(officer_stats[o]["hoan_thanh"] for o in sorted_officers),
                        "100.0%"
                    ]
                    self.pivot_tree.insert("", tk.END, values=tot_row, tags=("p_total",))

            elif "TỘI PHẠM" in mode_u or "LOẠI TỘI" in mode_u:
                # Chế độ 5: VKS Thụ Lý × Phân Loại Tội Phạm
                matrix, crime_cols, sorted_units = build_pivot_case_summary(filtered)
                cols = ["stt", "vks_thu_ly", "so_vu"] + [f"crime_{i}" for i in range(len(crime_cols))] + ["ty_le"]
                self.pivot_tree["columns"] = cols
                self.pivot_tree.heading("stt", text="STT")
                self.pivot_tree.column("stt", width=50, anchor=tk.CENTER)
                self.pivot_tree.heading("vks_thu_ly", text="Đơn vị VKS thụ lý")
                self.pivot_tree.column("vks_thu_ly", width=260, anchor=tk.W)
                self.pivot_tree.heading("so_vu", text="Tổng số vụ")
                self.pivot_tree.column("so_vu", width=90, anchor=tk.CENTER)
                for i, c in enumerate(crime_cols):
                    cid = f"crime_{i}"
                    self.pivot_tree.heading(cid, text=c)
                    self.pivot_tree.column(cid, width=120, anchor=tk.CENTER)
                self.pivot_tree.heading("ty_le", text="Tỷ lệ %")
                self.pivot_tree.column("ty_le", width=85, anchor=tk.CENTER)

                tot_all = len(filtered)
                for idx, u in enumerate(sorted_units, 1):
                    data = matrix[u]
                    pct = (data["total_cases"] / tot_all * 100) if tot_all > 0 else 0
                    row_vals = [idx, u, data["total_cases"]] + [data["crimes"].get(c, 0) or "" for c in crime_cols] + [f"{pct:.1f}%"]
                    row_tag = "p_even" if idx % 2 == 0 else "p_odd"
                    self.pivot_tree.insert("", tk.END, values=row_vals, tags=(row_tag,))

                if sorted_units:
                    tot_row = ["∑", "TỔNG CỘNG CHUNG", tot_all] + [sum(matrix[u]["crimes"].get(c, 0) for u in sorted_units) for c in crime_cols] + ["100.0%"]
                    self.pivot_tree.insert("", tk.END, values=tot_row, tags=("p_total",))

            else:
                # Chế độ 1 (MẶC ĐỊNH): VKS Thụ Lý × Giai Đoạn Tố Tụng
                matrix, active_stages, sorted_units = build_pivot_stage_summary(filtered)
                cols = ["stt", "vks_thu_ly"] + [f"stg_{i}" for i in range(len(active_stages))] + ["grand_total", "ty_le"]
                self.pivot_tree["columns"] = cols
                self.pivot_tree.heading("stt", text="STT")
                self.pivot_tree.column("stt", width=50, anchor=tk.CENTER)
                self.pivot_tree.heading("vks_thu_ly", text="Đơn vị VKS thụ lý (Row Labels)")
                self.pivot_tree.column("vks_thu_ly", width=280, anchor=tk.W)
                for i, s in enumerate(active_stages):
                    cid = f"stg_{i}"
                    self.pivot_tree.heading(cid, text=s)
                    self.pivot_tree.column(cid, width=120, anchor=tk.CENTER)
                self.pivot_tree.heading("grand_total", text="Grand Total")
                self.pivot_tree.column("grand_total", width=100, anchor=tk.CENTER)
                self.pivot_tree.heading("ty_le", text="Tỷ lệ %")
                self.pivot_tree.column("ty_le", width=85, anchor=tk.CENTER)

                tot_all = len(filtered)
                for idx, u in enumerate(sorted_units, 1):
                    st = matrix[u]
                    pct = (st["total_cases"] / tot_all * 100) if tot_all > 0 else 0
                    stg_vals = [st["stages"].get(s, 0) or "" for s in active_stages]
                    row_vals = [idx, u] + stg_vals + [st["total_cases"], f"{pct:.1f}%"]
                    row_tag = "p_even" if idx % 2 == 0 else "p_odd"
                    self.pivot_tree.insert("", tk.END, values=row_vals, tags=(row_tag,))

                if sorted_units:
                    tot_stgs = [sum(matrix[u]["stages"].get(s, 0) for u in sorted_units) for s in active_stages]
                    tot_row = ["∑", "Grand Total (TỔNG CỘNG CHUNG)"] + tot_stgs + [tot_all, "100.0%"]
                    self.pivot_tree.insert("", tk.END, values=tot_row, tags=("p_total",))
        else:
            matrix, stages, sorted_provinces = build_pivot_summary(filtered)
            pivot_cols = ["stt", "tinh_thanh", "so_vu", "tong_bi_can"] + [f"stg_{i}" for i in range(len(stages))] + ["ty_le"]
            self.pivot_tree["columns"] = pivot_cols
            self.pivot_tree.heading("stt", text="STT")
            self.pivot_tree.column("stt", width=50, anchor=tk.CENTER)
            self.pivot_tree.heading("tinh_thanh", text="Tỉnh / Thành phố")
            self.pivot_tree.column("tinh_thanh", width=190, anchor=tk.W)
            self.pivot_tree.heading("so_vu", text="Số vụ án")
            self.pivot_tree.column("so_vu", width=85, anchor=tk.CENTER)
            self.pivot_tree.heading("tong_bi_can", text="Tổng số bị can")
            self.pivot_tree.column("tong_bi_can", width=110, anchor=tk.CENTER)

            for i, s in enumerate(stages):
                col_id = f"stg_{i}"
                self.pivot_tree.heading(col_id, text=f"GĐ: {s}")
                self.pivot_tree.column(col_id, width=105, anchor=tk.CENTER)

            self.pivot_tree.heading("ty_le", text="Tỷ lệ %")
            self.pivot_tree.column("ty_le", width=85, anchor=tk.CENTER)

            for idx, prov in enumerate(sorted_provinces, start=1):
                data = matrix[prov]
                pct = (data["total_defendants"] / filter_defs * 100) if filter_defs > 0 else 0
                row_vals = [
                    idx,
                    prov,
                    len(data["cases"]),
                    data["total_defendants"],
                ] + [data["stages"].get(s, 0) for s in stages] + [f"{pct:.1f}%"]
                row_tag = "p_even" if idx % 2 == 0 else "p_odd"
                self.pivot_tree.insert("", tk.END, values=row_vals, tags=(row_tag,))

            if sorted_provinces:
                total_all_cases = len(set(r.get("ma_vu_an", "").strip() for r in filtered if r.get("ma_vu_an", "").strip()))
                tot_row = [
                    "∑",
                    "TỔNG CỘNG CHUNG",
                    total_all_cases,
                    filter_defs,
                ] + [sum(matrix[p]["stages"].get(s, 0) for p in sorted_provinces) for s in stages] + ["100.0%"]
                self.pivot_tree.insert("", tk.END, values=tot_row, tags=("p_total",))

        if hasattr(self, "lbl_pivot_desc") and hasattr(self, "cbo_pivot_mode"):
            mode_s = self.cbo_pivot_mode.get()
            mode_u = mode_s.upper()
            if "RỦI RO" in mode_u or "THỜI HIỆU" in mode_u:
                self.lbl_pivot_desc.config(text="Cảnh báo mức độ nguy cơ: Đỏ đậm = Hết thời hiệu (Đ.230 BLTTHS); Đỏ tươi = Còn ≤ 7 ngày; Cam = 8-15 ngày.", fg="#991B1B")
            elif "TIẾN ĐỘ" in mode_u or "GIAI ĐOẠN" in mode_u or "KẾT QUẢ" in mode_u:
                self.lbl_pivot_desc.config(text="Ma trận đối soát phân bổ khối lượng án theo từng Giai đoạn tố tụng (Điều tra, Truy tố, Sơ thẩm) giữa các đơn vị.", fg="#1E40AF")
            elif "PHÂN CÔNG" in mode_u or "KIỂM SÁT VIÊN" in mode_u or "KSV" in mode_u:
                self.lbl_pivot_desc.config(text="Giám sát trách nhiệm cá nhân, khối lượng công việc và tình hình xử lý án quá hạn của từng Kiểm sát viên.", fg="#047857")
            elif "TOÀN DIỆN" in mode_u or "360" in mode_s:
                self.lbl_pivot_desc.config(text="Báo cáo điều hành tổng quan 360° đa chiều gửi Lãnh đạo Viện (Giai đoạn, Thời hiệu, Tội danh, Bị can).", fg="#B45309")
            elif "TỘI PHẠM" in mode_u or "LOẠI TỘI" in mode_u:
                self.lbl_pivot_desc.config(text="Thống kê cơ cấu phân loại tính chất tội phạm (Ít nghiêm trọng, Nghiêm trọng, Rất NT, Đặc biệt NT).", fg="#166534")
            else:
                self.lbl_pivot_desc.config(text="Tự động trích xuất và tổng hợp số liệu ma trận Pivot phục vụ báo cáo lãnh đạo Viện.", fg="#475569")

    def on_open_excel(self):
        if openpyxl is None:
            messagebox.showerror("Thiếu thư viện", "Vui lòng cài đặt openpyxl: pip install openpyxl")
            return
        fpaths = filedialog.askopenfilenames(
            title="Chọn file Excel copy từ QLAHS (có thể chọn 1 hoặc nhiều file cùng lúc)",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if not fpaths:
            return

        fpaths = list(fpaths)
        n_files = len(fpaths)

        if n_files > 1:
            self._execute_reconcile_flow(fpaths)
            return

        fpath = fpaths[0]
        fname = os.path.basename(fpath)

        def _worker(update_cb):
            return load_excel_and_parse(fpath, update_cb)

        def _on_success(res):
            fmt, records = res
            self.data_format = fmt
            if fmt in ("CASE", "CASE_REPORT", "STANDARDIZED_CASE"):
                for r in records:
                    normalize_case_stage_and_status(r, filename=fname)
            self.records = records
            self.refresh_views()

            mode = "TIN_BAO" if fmt in ("TIN_BAO", "STANDARDIZED_TIN_BAO") else ("DEFENDANT" if fmt == "DEFENDANT" else "CASE")
            file_stats = [{
                "fname": fname,
                "total": len(records),
                "new": len(records),
                "matched": 0,
            }]
            self.last_reconcile_result = (mode, records, 0, file_stats, [(fname, fmt, records)])

            if fmt in ("TIN_BAO", "STANDARDIZED_TIN_BAO"):
                unit_label = "nguồn tin báo về tội phạm (36 cột chuẩn hóa)"
            elif fmt == "DEFENDANT":
                unit_label = "hồ sơ bị can / đối tượng truy nã, đình nã (12 cột chuẩn hóa)"
            elif fmt in ("CASE", "CASE_REPORT", "STANDARDIZED_CASE"):
                unit_label = "vụ án (27 cột chuẩn hóa)"
            else:
                unit_label = "hồ sơ bị can (12 cột chuẩn hóa)"

            self.lbl_status.config(
                text=f"✓ Đã chuẩn hóa thành công {len(records):,} {unit_label} từ tập tin {fname}",
                fg="#166534"
            )

            # Mở Dashboard báo cáo kết quả chuẩn hóa & đối chiếu trực quan thay thế messagebox
            ReconciliationDashboardDialog(
                parent=self,
                mode=mode,
                records=records,
                match_n=0,
                file_stats=file_stats,
                on_export_cb=self.on_export_excel,
                on_view_table_cb=lambda: self.notebook.select(0),
                on_view_pivot_cb=lambda: self.notebook.select(1)
            )

        self.run_async_task(
            title="Đang nạp file Excel QLAHS",
            loading_msg=f"Hệ thống đang mở và xử lý tập tin {fname}...",
            worker_fn=_worker,
            on_success=_on_success
        )

    def on_paste_clipboard(self):
        try:
            clip = self.clipboard_get()
        except Exception:
            messagebox.showwarning("Clipboard trống", "Không tìm thấy dữ liệu trong Clipboard.")
            return

        if not clip.strip():
            messagebox.showwarning("Clipboard trống", "Dữ liệu Clipboard rỗng.")
            return

        def _worker(update_cb):
            update_cb("Đang phân tích cấu trúc dữ liệu từ Clipboard...", "Tách dòng và cột...")
            return parse_clipboard_data(clip, update_cb)

        def _on_success(res):
            fmt, records = res
            self.data_format = fmt
            if fmt in ("CASE", "CASE_REPORT"):
                for r in records:
                    normalize_case_stage_and_status(r, filename="Clipboard")
            self.records = records
            self.refresh_views()
            mode = "TIN_BAO" if fmt == "TIN_BAO" else ("DEFENDANT" if fmt == "DEFENDANT" else "CASE")
            file_stats = [{
                "fname": "Dữ liệu Clipboard (Ctrl+V)",
                "total": len(records),
                "new": len(records),
                "matched": 0,
            }]
            self.last_reconcile_result = (mode, records, 0, file_stats, [("Clipboard", fmt, records)])

            if fmt in ("TIN_BAO", "STANDARDIZED_TIN_BAO"):
                unit_name = "nguồn tin báo & tố giác tội phạm (36 cột chuẩn hóa)"
            elif fmt == "DEFENDANT":
                unit_name = "hồ sơ bị can / đối tượng truy nã, đình nã (12 cột chuẩn hóa)"
            elif fmt in ("CASE", "CASE_REPORT", "STANDARDIZED_CASE"):
                unit_name = "vụ án (27 cột chuẩn hóa)"
            else:
                unit_name = "hồ sơ bị can (12 cột chuẩn hóa)"

            self.lbl_status.config(
                text=f"✓ Đã chuẩn hóa thành công {len(records):,} {unit_name} từ Clipboard",
                fg="#166534"
            )

            ReconciliationDashboardDialog(
                parent=self,
                mode=mode,
                records=records,
                match_n=0,
                file_stats=file_stats,
                on_export_cb=self.on_export_excel,
                on_view_table_cb=lambda: self.notebook.select(0),
                on_view_pivot_cb=lambda: self.notebook.select(1)
            )

        self.run_async_task(
            title="Đang xử lý dữ liệu Clipboard",
            loading_msg="Hệ thống đang phân tích dữ liệu dán từ Clipboard...",
            worker_fn=_worker,
            on_success=_on_success
        )

    def show_reconciliation_dashboard(self):
        """Mở lại Dashboard kết quả đối chiếu gần nhất."""
        if hasattr(self, "last_reconcile_result") and self.last_reconcile_result:
            mode, merged, match_n, file_stats, _ = self.last_reconcile_result
            ReconciliationDashboardDialog(
                parent=self,
                mode=mode,
                records=merged,
                match_n=match_n,
                file_stats=file_stats,
                on_export_cb=self.on_export_excel,
                on_view_table_cb=lambda: self.notebook.select(0),
                on_view_pivot_cb=lambda: self.notebook.select(1)
            )
            return

        if self.records:
            src_map = {}
            for r in self.records:
                s = r.get("nguon_du_lieu") or "Dữ liệu hiện hành"
                src_map[s] = src_map.get(s, 0) + 1
            file_stats = [
                {"fname": fn, "total": cnt, "new": cnt, "matched": 0}
                for fn, cnt in src_map.items()
            ]
            mode = "TIN_BAO" if self.data_format == "TIN_BAO" else "CASE"
            ReconciliationDashboardDialog(
                parent=self,
                mode=mode,
                records=self.records,
                match_n=0,
                file_stats=file_stats,
                on_export_cb=self.on_export_excel,
                on_view_table_cb=lambda: self.notebook.select(0),
                on_view_pivot_cb=lambda: self.notebook.select(1)
            )
            return

        messagebox.showinfo(
            "Dashboard Đối Chiếu",
            "Chưa có dữ liệu nào được nạp vào phần mềm.\n\n"
            "Vui lòng nhấn nút '📂 1. Chọn File' hoặc '⚡ 4. Đối chiếu & Hợp nhất Nhiều File' trước."
        )

    def on_export_excel(self):
        if not self.records:
            messagebox.showwarning("Chưa có dữ liệu", "Chưa có dữ liệu để xuất ra Excel!")
            return

        is_enriched = (self.data_format == "CASE_ENRICHED")
        is_tin_bao = (self.data_format in ("TIN_BAO", "STANDARDIZED_TIN_BAO"))
        is_reconcile = bool(self.records and any(r.get("nguon_du_lieu") for r in self.records)) and not is_enriched
        is_case = (self.data_format in ("CASE", "CASE_REPORT", "CASE_ENRICHED", "STANDARDIZED_CASE") or (is_reconcile and not is_tin_bao))

        time_str = datetime.now().strftime("%d%m%Y_%Hh%M")
        if is_enriched:
            init_file = f"Ket_Qua_Lien_Thong_Vu_An_Va_Bi_Can_{time_str}.xlsx"
        elif is_tin_bao and is_reconcile:
            init_file = f"Ket_Qua_Ra_Soat_Hop_Nhat_Tin_Bao_{time_str}.xlsx"
        elif is_tin_bao:
            init_file = f"Ket_Qua_Chuan_Hoa_Tin_Bao_{time_str}.xlsx"
        elif is_reconcile:
            init_file = f"Ket_Qua_Ra_Soat_Doi_Chieu_QLAHS_{time_str}.xlsx"
        elif is_case:
            init_file = f"Ket_Qua_Chuan_Hoa_So_Thu_Ly_Vu_An_{time_str}.xlsx"
        else:
            init_file = f"Ket_Qua_Chuan_Hoa_Bi_Can_Truy_Na_{time_str}.xlsx"

        out_path = filedialog.asksaveasfilename(
            title="Lưu file Excel chuẩn hóa và Pivot Table",
            defaultextension=".xlsx",
            initialfile=init_file,
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not out_path:
            return

        def _worker(update_cb):
            if is_enriched:
                l_cnt = len(getattr(self, "linked_defendants", []))
                u_cnt = len(getattr(self, "unlinked_defendants", []))
                update_cb(f"Đang xuất file liên thông làm giàu {len(self.records):,} vụ án...", f"Khớp nối {l_cnt:,} bị can, tạo 5 Sheet chuyên nghiệp...")
                export_case_enriched_to_excel(self.records, getattr(self, "linked_defendants", []), getattr(self, "unlinked_defendants", []), out_path)
            elif is_tin_bao:
                sheet_msg = "Tạo 3 Sheet: Dữ liệu hợp nhất 36 cột, Báo cáo tiến độ & Pivot Table..." if is_reconcile else "Tạo 3 Sheet: Dữ liệu chuẩn 35 cột, Báo cáo tiến độ & Pivot Table..."
                update_cb(f"Đang xuất {len(self.records):,} tin báo tội phạm...", sheet_msg)
                export_tin_bao_to_excel(self.records, out_path)
            elif is_reconcile:
                update_cb(f"Đang xuất file hợp nhất đối chiếu {len(self.records):,} vụ án...", "Tạo 5 Sheet: Tổng quan Pivot Lãnh đạo, Pivot Giai đoạn, Dữ liệu chuẩn...")
                export_reconciliation_excel(self.records, out_path)
            elif is_case:
                update_cb(f"Đang định dạng {len(self.records):,} vụ án...", "Tạo 4 Sheet: Tổng quan Pivot Lãnh đạo, Pivot Giai đoạn, Sổ vụ án...")
                export_case_records_to_excel(self.records, out_path)
            else:
                update_cb(f"Đang định dạng {len(self.records):,} dòng bị can...", "Đang tạo Sheet Danh sách & Pivot Table...")
                export_records_to_excel(self.records, out_path)
            return len(self.records)

        def _on_success(count):
            if is_enriched:
                l_cnt = len(getattr(self, "linked_defendants", []))
                u_cnt = len(getattr(self, "unlinked_defendants", []))
                unit_name = f"vụ án đã liên thông & làm giàu ({l_cnt:,} bị can đã liên thông, {u_cnt:,} bị can chờ xác minh PM cũ - 5 Sheet chuyên nghiệp)"
            elif is_tin_bao:
                unit_name = "nguồn tin báo về tội phạm (36 cột hợp nhất, 3 Sheet)" if is_reconcile else "nguồn tin báo về tội phạm (35 cột, 3 Sheet)"
            elif is_case:
                unit_name = "vụ án (29 cột hợp nhất, 5 Sheet kèm Pivot Lãnh đạo)" if is_reconcile else "vụ án (26 cột, 4 Sheet kèm Pivot Lãnh đạo)"
            else:
                unit_name = "hồ sơ bị can / đối tượng truy nã & đình nã (12 cột chuẩn hóa, 2 Sheet)"
            messagebox.showinfo(
                "Xuất file Excel thành công",
                f"Đã kết xuất thành công {count:,} {unit_name} và Bảng Pivot Table sang file:\n{out_path}"
            )

        self.run_async_task(
            title="Đang xuất file Excel",
            loading_msg="Đang định dạng các Sheet Báo cáo Tổng quan Pivot & Dữ liệu chuẩn...",
            worker_fn=_worker,
            on_success=_on_success
        )


    def on_clear(self):
        self.records = []
        self.refresh_views()

    def _execute_reconcile_flow(self, fpaths):
        """Thực thi luồng nạp, đối chiếu và hợp nhất nhiều file QLAHS."""
        n_files = len(fpaths)
        self.lbl_status.config(
            text=f"⏳ Đang đọc và đối chiếu {n_files} file QLAHS...",
            fg="#D97706"
        )
        self.update_idletasks()

        def _worker(update_cb):
            parsed_files = []
            for i, fp in enumerate(fpaths):
                fname = os.path.basename(fp)
                update_cb(
                    f"Đang đọc tập tin ({i+1}/{n_files}): {fname}...",
                    f"Phân tích cấu trúc dữ liệu..."
                )
                fmt, recs = load_excel_and_parse(fp, update_cb)
                parsed_files.append((fname, fmt, recs))

            # Tự động nhận diện loại dữ liệu đối chiếu: Tin Báo hay Vụ Án hay Liên thông Vụ án - Bị can
            case_files = [p for p in parsed_files if p[1] in ("CASE", "CASE_REPORT", "STANDARDIZED_CASE")]
            def_files = [p for p in parsed_files if p[1] == "DEFENDANT"]
            tin_bao_files = [p for p in parsed_files if p[1] in ("TIN_BAO", "STANDARDIZED_TIN_BAO")]

            if case_files and def_files and not tin_bao_files:
                update_cb(
                    f"Đang liên thông & làm giàu dữ liệu Sổ vụ án ⟷ Danh sách bị can...",
                    f"{len(case_files)} file Vụ án ({sum(len(r) for _, _, r in case_files):,} vụ) & {len(def_files)} file Bị can ({sum(len(r) for _, _, r in def_files):,} bị can)"
                )
                cases, match_n, file_stats, linked_defs, unlinked_defs = reconcile_and_link_cases_defendants(
                    case_files, def_files, update_cb
                )
                update_cb(
                    f"Đang chuẩn bị hiển thị dữ liệu liên thông...",
                    f"Hoàn tất: {match_n} vụ án được làm giàu, {len(linked_defs)} bị can đã liên thông, {len(unlinked_defs)} bị can chờ xác minh"
                )
                return "CASE_ENRICHED", cases, match_n, file_stats, parsed_files, linked_defs, unlinked_defs

            tin_bao_count = sum(1 for _, fmt, _ in parsed_files if fmt == "TIN_BAO")
            is_tin_bao_mode = (tin_bao_count >= len(parsed_files) / 2)

            if is_tin_bao_mode:
                update_cb(
                    f"Đang so khớp chéo & bù đắp thông tin {n_files} tập tin Tin báo...",
                    f"Tổng cộng {sum(len(r) for _, _, r in parsed_files):,} tin báo đầu vào"
                )
                merged, match_n, file_stats = reconcile_and_merge_multiple_tin_bao_lists(parsed_files, update_cb)
                update_cb(
                    f"Đang chuẩn bị hiển thị danh sách 36 cột hợp nhất...",
                    f"Tổng hợp trọn vẹn {len(merged):,} tin báo"
                )
                return "TIN_BAO", merged, match_n, file_stats, parsed_files
            else:
                update_cb(
                    f"Đang so khớp chéo & bù đắp thông tin {n_files} tập tin Vụ án...",
                    f"Tổng cộng {sum(len(r) for _, _, r in parsed_files):,} vụ án đầu vào"
                )
                merged, match_n, file_stats = reconcile_and_merge_multiple_case_lists(parsed_files, update_cb)
                update_cb(
                    f"Đang chuẩn bị hiển thị danh sách 27 cột...",
                    f"Tổng hợp trọn vẹn {len(merged):,} vụ án"
                )
                return "CASE", merged, match_n, file_stats, parsed_files

        def _on_success(res):
            if len(res) == 7 and res[0] == "CASE_ENRICHED":
                mode, merged, match_n, file_stats, parsed_files, linked_defs, unlinked_defs = res
                self.linked_defendants = linked_defs
                self.unlinked_defendants = unlinked_defs
                self.linked_defs_count = len(linked_defs)
                self.unlinked_defs_count = len(unlinked_defs)
            else:
                mode, merged, match_n, file_stats, parsed_files = res
                self.linked_defendants = []
                self.unlinked_defendants = []
                self.linked_defs_count = 0
                self.unlinked_defs_count = 0

            self.data_format = mode
            if mode in ("CASE", "CASE_REPORT", "CASE_ENRICHED"):
                for r in merged:
                    normalize_case_stage_and_status(r)
            self.records = merged
            self.last_reconcile_result = res
            self.refresh_views()

            if mode == "CASE_ENRICHED":
                self.lbl_status.config(
                    text=f"✓ Đã liên thông thành công {len(merged):,} vụ án ({match_n} vụ được làm giàu bị can, {len(linked_defs):,} bị can đã khớp, {len(unlinked_defs):,} bị can chờ xác minh PM cũ)",
                    fg="#166534"
                )
            else:
                unit_label = "tin báo về tội phạm" if mode == "TIN_BAO" else "vụ án"
                self.lbl_status.config(
                    text=f"✓ Đã hợp nhất thành công {len(merged):,} {unit_label} từ {len(fpaths)} file QLAHS",
                    fg="#166534"
                )

            # Mở Dashboard trực quan chuyên nghiệp
            ReconciliationDashboardDialog(
                parent=self,
                mode=mode,
                records=merged,
                match_n=match_n,
                file_stats=file_stats,
                on_export_cb=self.on_export_excel,
                on_view_table_cb=lambda: self.notebook.select(0),
                on_view_pivot_cb=lambda: self.notebook.select(1)
            )

        self.run_async_task(
            title=f"Đang đối chiếu & hợp nhất {n_files} File QLAHS",
            loading_msg=f"Hệ thống đang mở và so khớp {n_files} tập tin Excel...",
            worker_fn=_worker,
            on_success=_on_success,
            on_error=lambda err: messagebox.showerror("Lỗi đối chiếu", f"Không thể đối chiếu các file:\n{err}")
        )

    def on_reconcile_multiple_files(self):
        """Mở nhiều file Excel QLAHS cùng lúc (không giới hạn 2 file),
        tự động phân tích cấu trúc, nhận diện xem là Vụ án hay Tin báo,
        rà soát đối chiếu chéo, bù đắp làm giàu dữ liệu và hợp nhất theo đúng định dạng tương ứng."""
        if openpyxl is None:
            messagebox.showerror("Thiếu thư viện", "Vui lòng cài đặt openpyxl: pip install openpyxl")
            return

        fpaths = filedialog.askopenfilenames(
            title="Chọn các file Excel QLAHS cần đối chiếu & hợp nhất (Giữ Ctrl hoặc Shift để chọn nhiều file)",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if not fpaths:
            return

        fpaths = list(fpaths)
        n_files = len(fpaths)

        if n_files == 1:
            ans = messagebox.askyesno(
                "Đang chọn 1 file",
                f"Bạn vừa chọn 1 file:\n{os.path.basename(fpaths[0])}\n\n"
                "Bạn có muốn chọn thêm các file khác để đối chiếu chéo không?\n"
                "(Bấm 'Yes' để chọn thêm file, bấm 'No' để nạp luôn file này)"
            )
            if ans:
                fpaths2 = filedialog.askopenfilenames(
                    title="Chọn thêm các file Excel QLAHS khác để đối chiếu",
                    filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
                )
                if fpaths2:
                    for fp in fpaths2:
                        if fp not in fpaths:
                            fpaths.append(fp)

        self._execute_reconcile_flow(fpaths)

    # Alias tương thích ngược
    on_reconcile_two_files = on_reconcile_multiple_files

    def on_diff_analysis(self):
        """Mở hộp thoại so sánh biến động giữa 2 kỳ báo cáo (Diff Analysis)."""
        DiffAnalysisDialog(self, current_records=self.records, current_format=self.data_format)

    def on_report_summary(self):
        """Mở hộp thoại trích yếu báo cáo nghiệp vụ tự động."""
        if not self.records:
            messagebox.showwarning(
                "Chưa có dữ liệu",
                "Vui lòng nạp ít nhất một danh sách dữ liệu QLAHS (mở file hoặc dán clipboard) trước khi xuất báo cáo trích yếu.",
                parent=self
            )
            return
        ReportSummaryDialog(self, records=self.records, data_format=self.data_format, config=CURRENT_CONFIG)

    def on_open_settings(self):
        """Mở bảng cấu hình quy tắc mở (Settings)."""
        SettingsDialog(self, config=CURRENT_CONFIG, on_save_cb=self._on_settings_saved)

    def _on_settings_saved(self, new_cfg):
        """Áp dụng các cấu hình mới trực tiếp lên giao diện đang chạy."""
        agency = new_cfg.get("agency_name", DEFAULT_APP_CONFIG["agency_name"])
        author = new_cfg.get("author_title", DEFAULT_APP_CONFIG["author_title"])
        self.title(f"{agency} - Hệ thống Chuẩn hóa & Đối chiếu QLAHS (Version 1.0 Offline)")
        if hasattr(self, "lbl_footer_author"):
            self.lbl_footer_author.config(text=f"Phát triển bởi: {author}")
        # Cập nhật lại views nếu đang có records
        if self.records:
            self.refresh_views()


if __name__ == "__main__":
    app = QLAHSApp()
    app.mainloop()
