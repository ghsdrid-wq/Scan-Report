# Scan Report Tool

<img width="722" height="752" alt="image" src="https://github.com/user-attachments/assets/94137768-ac3f-4521-88fa-36a0e42c0a77" />

Python desktop application สำหรับประมวลผลข้อมูล Scan Report และสร้างรายงานสรุปแบบ Pivot Table พร้อมระบบ Heatmap, Shift Summary, Category Coloring และ Excel Dashboard Export

โปรแกรมถูกออกแบบสำหรับ warehouse operation และ logistics monitoring เพื่อวิเคราะห์ปริมาณการสแกนรายสถานีตามช่วงเวลา

รองรับ:

- Scan Report Processing
- Time Slot Analysis
- Pivot Summary
- Shift Calculation
- Heatmap Visualization
- Excel Dashboard Export
- Category Coloring
- GUI Monitoring
- Real-time Progress Tracking

---

# Features

- Excel scan report processing
- Auto time slot generation
- Station pivot summary
- Shift A / Shift B calculation
- Grand total calculation
- Heatmap visualization
- Category color grouping
- Auto Excel formatting
- Freeze pane + filter
- Real-time progress tracking
- ETA estimation
- Config persistence
- Time range filtering
- Auto column width
- Background threading
- Windows desktop GUI

---

# Application Overview

โปรแกรมนี้ใช้สำหรับ:

1. อ่านข้อมูล Scan Report
2. แปลงข้อมูลเป็น time slot
3. สร้าง pivot summary ราย station
4. คำนวณยอดแต่ละ shift
5. สร้าง dashboard report
6. Export Excel พร้อม styling และ heatmap

เหมาะสำหรับ:

- Warehouse Monitoring
- Logistics Operation
- Loading Station Analysis
- Scan Throughput Report
- Shift Performance Tracking
- Operation Dashboard

---

# Tech Stack

- Python
- Tkinter
- Pandas
- OpenPyXL
- Threading
- ConfigParser

---

# Project Structure

```text
project/
│
├── main.py
├── config.ini
│
└── output/
```

---

# Scan Report Workflow

```text
Load Excel File
    ↓
Prepare Scan Data
    ↓
Generate Time Slot
    ↓
Pivot by Station
    ↓
Calculate Shift Summary
    ↓
Generate Total Row
    ↓
Apply Excel Style
    ↓
Apply Heatmap
    ↓
Export Dashboard Report
```

---

# Input Data Processing

ระบบจะอ่าน:

- scan_time
- station

จาก column index ภายในไฟล์ Excel

และแปลง:

```python
scan_time -> datetime
```

---

# Time Slot Logic

ระบบจะสร้าง time slot อัตโนมัติแบบ:

```text
12.00-13.00
13.00-14.00
14.00-15.00
```

โดย:

- ชั่วโมงก่อน 12:00 จะถูก shift +24
- รองรับ operation ที่ทำงานข้ามวัน

---

# Pivot Report

ระบบจะสร้าง pivot table:

| Station | Time Slot | Count |
|---|---|---|
| Station A | 12.00-13.00 | 125 |
| Station A | 13.00-14.00 | 148 |

โดยใช้:

```python
pandas.groupby().size().unstack()
```

---

# Time Range Filter

ผู้ใช้สามารถเลือก:

- Start Hour
- End Hour

ผ่าน GUI

ตัวอย่าง:

```text
12:00 → 18:00
```

ระบบจะ filter เฉพาะช่วงเวลาที่เลือก

---

# Shift Calculation

ระบบแบ่ง:

## Shift A

ช่วงเวลา:

```text
12:00 - 00:00
```

---

## Shift B

ช่วงเวลา:

```text
00:00 - 12:00
```

---

# Generated Summary Columns

ระบบจะสร้าง:

| Column |
|---|
| Total Shift A |
| Total Shift B |
| Grand Total |

พร้อม auto calculation

---

# Total Summary Row

ระบบจะเพิ่ม:

```text
TOTAL ALL STATIONS
```

เพื่อสรุปยอดรวมทั้งหมด

---

# Station Category Coloring

รองรับ category color อัตโนมัติ

## Gateway

```text
*_GW
```

สี:

```text
Blue
```

---

## DWS

```text
D_*
```

สี:

```text
Light Blue
```

---

## Numeric Station

```text
0-9*
```

สี:

```text
Soft Blue
```

---

# Heatmap System

รองรับ conditional formatting heatmap

ใช้:

```python
ColorScaleRule
```

สำหรับ:

- low value = red
- medium value = yellow
- high value = green

---

# Heatmap Safety System

ระบบใช้:

```python
apply_heatmap_safe()
```

เพื่อ:

- ตรวจสอบ column mapping
- ป้องกัน invalid range
- skip auto หาก column ไม่ครบ

---

# Excel Styling

ระบบใช้:

```python
openpyxl
```

สำหรับ:

- header styling
- category coloring
- total row formatting
- border styling
- alignment
- auto width
- freeze pane
- auto filter

---

# GUI Features

โปรแกรมใช้:

```python
Tkinter + ttk
```

ประกอบด้วย:

- Input Browser
- Progress Bar
- ETA Display
- Real-time Log
- Export Settings
- Heatmap Toggle
- Category Toggle
- Process / Cancel Button

---

# Progress Tracking

GUI แสดง:

- progress percentage
- ETA estimation
- current step
- processing log

ETA คำนวณจาก:

```python
(elapsed / progress) * remaining
```

---

# Config System

ใช้:

```text
config.ini
```

เก็บ:

- last input path
- last output path
- time range
- heatmap setting
- category setting

---

# Example Config

```ini
[PATH]
input=C:/INPUT/report.xlsx
output=C:/OUTPUT

[TIME]
start=12:00
end=12:00

[OPTION]
heatmap=True
category=True
```

---

# Background Processing

โปรแกรมใช้:

```python
threading.Thread()
```

เพื่อ:

- ป้องกัน GUI freeze
- รองรับ large dataset
- smooth progress update
- responsive UI

---

# Queue Communication System

ใช้:

```python
queue.Queue()
```

สำหรับ:

- log update
- progress update
- finish notification
- GUI synchronization

---

# Excel Output Example

```text
สแกนบรรจุขึ้นรถ_2026-05-17.xlsx
```

ภายในประกอบด้วย:

- Station summary
- Time slot report
- Shift summary
- Heatmap visualization
- Grand total summary

---

# Installation

## 1. Clone Repository

```bash
git clone https://github.com/yourname/scan-report-tool.git
```

---

## 2. Install Dependencies

```bash
pip install pandas openpyxl
```

---

## 3. Run Application

```bash
python main.py
```

---

# Build EXE

ใช้ PyInstaller:

```bash
pyinstaller --onefile --windowed main.py
```

หรือ:

```bash
pyinstaller --noconsole --onefile main.py
```

---

# Error Handling

ระบบรองรับ:

- Invalid Excel file
- Invalid datetime
- Missing column
- Save cancelled
- Heatmap column mismatch
- Invalid time range
- Pivot generation error
- Export failure

---

# User Experience Features

- Real-time processing log
- Heatmap toggle
- Category color toggle
- Auto remember path
- Auto remember settings
- ETA display
- Styled dashboard output
- Freeze pane support
- Auto filter support

---

# Future Improvements

- Multi-sheet dashboard
- Chart visualization
- PDF export
- Auto scheduler
- Database integration
- Live dashboard mode
- CSV import support
- Drag & drop file support
- Multi-report merge
- Station KPI analytics

---

# License

MIT License

---

# Author

Developed for warehouse scan throughput monitoring and logistics operation reporting workflow.


