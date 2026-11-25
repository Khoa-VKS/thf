@echo off
chcp 65001 >nul
title Dong Goi EXE - VKSND TP. HỒ CHÍ MINH - He Thong NTQLAHS
echo Dang dong goi thanh file .EXE chay doc lap khong can cai Python...
pyinstaller --noconsole --onefile --name "He_Thong_NTQLAHS_VKSND_TPHCM" Phan_Mem_Chuan_Hoa_QLAHS_Desktop.py
echo Hoan tat! File .exe nam trong thu muc dist\He_Thong_NTQLAHS_VKSND_TPHCM.exe
pause
