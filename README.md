# 📊 Sales Report Automation

Python script ที่ช่วยรวมข้อมูลยอดขายจากหลายสาขา
และสร้างรายงาน Excel อัตโนมัติพร้อมจัดรูปแบบ

## ✨ Features
- 📂 อ่านไฟล์ Excel หลายสาขาพร้อมกันอัตโนมัติ
- 📊 สรุปยอดขายแยกตามสาขา
- 🏆 หา Top 20 สินค้าขายดีที่สุด
- 🎨 จัดรูปแบบ Excel สวยงาม พร้อม Freeze Header
- 📅 ตั้งชื่อไฟล์ตามวันที่อัตโนมัติ

## 🛠️ Tech Stack
- Python
- pandas
- openpyxl

## ⚙️ Installation
pip install pandas openpyxl

## 🚀 วิธีใช้
1. วางไฟล์ Excel ทุกสาขาในโฟลเดอร์ `input/`
2. รัน `python daily_report.py`
3. รับไฟล์รายงานในโฟลเดอร์ `output/`

## โครงสร้างโฟลเดอร์
daily_report/
├── input/
├── output/
└── daily_report.py
