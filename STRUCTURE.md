# STRUCTURE — Scan-Report (Scan Report Tool)

> ⚠️ **กฎการดูแลไฟล์นี้ (สำคัญ)**
> ทุกครั้งที่แก้ไขโค้ดใน repo นี้ — เพิ่ม/ลบ/ย้ายไฟล์, เปลี่ยน logic pivot/time slot/shift, เปลี่ยนคอลัมน์ที่อ้างอิง (index), เปลี่ยน style/heatmap, หรือเปลี่ยน flow — **ต้องอัปเดต STRUCTURE.md นี้ให้ตรงกับโค้ดเสมอ**

## ภาพรวม
แอป GUI (**Tkinter**) ชื่อ *"Scan Report Processor"* (V1.1, by Siwamon Singtan, IT KKN) — อ่านไฟล์ Excel log การสแกน → **pivot จำนวนสแกนเป็นตาราง Station × ช่วงเวลา (รายชั่วโมง)** → รวม Shift A/B + Grand Total → จัด style/สี/heatmap → export Excel

## วิธีรัน / Entry point
- รัน: `python main.py` → คลาส `App` (Tkinter, ทำงานใน thread + queue)

## โครงสร้างไฟล์
| ไฟล์ | หน้าที่ |
|------|---------|
| `main.py` | ทั้งโปรแกรม — UI, `run_process()` (pivot logic), `ask_save()` (เขียน + จัด style ด้วย openpyxl), `apply_heatmap_safe()` |

## Logic ใน run_process() (สำคัญ)
- **คอลัมน์อ้างอิงด้วย index:** scan_time = col 3, Station = col 5
- `get_time_slot()` — แปลงชั่วโมงเป็นช่วง `HH.00-HH.00` (ชั่วโมง < 12 บวก 24 เพื่อเรียงข้ามวัน)
- pivot: `groupby(['Station','time_slot']).size().unstack()`
- เรียง Station ด้วย `sort_key`: ลงท้าย `_GW` → ขึ้นต้น `D_` → ขึ้นต้นตัวเลข → อื่นๆ
- Shift A = ชั่วโมง ≥ 12 หรือ 0, Shift B = ที่เหลือ; เพิ่ม `Total Shift A/B`, `Grand Total`, แถว TOTAL
- เลือกช่วงเวลาที่แสดงตาม Time Range (start→end)

## ฟังก์ชันสำคัญ
- `apply_heatmap_safe(ws, enable)` — ColorScaleRule (3 ระดับ แดง-เหลือง-เขียว) บนโซน A/B/C ตามตำแหน่งคอลัมน์ Shift
- `ask_save()` — เขียนไฟล์, header fill, border, number format, freeze A2, auto-filter, category color, auto width

## Config (`config.ini`)
- `[PATH]`: `input`, `output`
- `[TIME]`: `start`, `end`
- `[OPTION]`: `heatmap`, `category` (เปิด/ปิดการลงสี)

## Dependencies
- `pandas`, `openpyxl` (Font/Fill/Border/ColorScaleRule), `tkinter`

## ข้อควรระวัง
- อ้างอิงคอลัมน์ด้วย index (3, 5) → ฟอร์แมต log ต้องตรงตำแหน่ง
- ชื่อไฟล์ผล default = `สแกนบรรจุขึ้นรถ_<วันที่ในไฟล์>.xlsx`
