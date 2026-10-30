# ระบบบริหารจัดการฟาร์มไก่ไข่ (Egg Farm Management System)
**เวอร์ชันลดจำนวนไฟล์ (Minimal Files Architecture)**

โครงสร้างได้รับการปรับลดจำนวนไฟล์ให้เหลือเพียง **2 ไฟล์หลัก** (พร้อมไฟล์ `appsscript.json`) ทำให้ติดตั้งและคัดลอกลงใน Google Apps Script Editor ได้ง่ายและรวดเร็วที่สุด:

1. **`Code.gs`**: รวมการทำงานฝั่ง Backend ทั้งหมด (Config, Database, CRUD Transactions, Dashboard Analytics, File Service Google Drive, Report Service A4, Camera Service และ doGet)
2. **`Index.html`**: รวมการทำงานฝั่ง Frontend ทั้งหมด (HTML Markup, ธีมสีสไตล์ CSS ใน `<style>`, และ JavaScript Client-side ใน `<script>`)
3. **`appsscript.json`**: ไฟล์ Manifest กำหนดค่า TimeZone (`Asia/Bangkok`) และ OAuth Scopes

---

## 1. ข้อมูลการเชื่อมต่อและการกำหนดค่า (Configuration)

- **GOOGLE_SHEET_ID**: `1AmRxzTRyd4-EerC4V9ISIuM1wSe2qNU_u22KkbTrLoQ`
- **GOOGLE_DRIVE_FOLDER_ID**: `13faPNclR-u_n4fd7wfF5fCqPHUhwczwn`

---

## 2. วิธีติดตั้งบน Google Apps Script ใน 3 นาที

1. เปิด [script.google.com](https://script.google.com) แล้วกด **"โปรเจกต์ใหม่" (New project)**
2. ไปที่ **การตั้งค่าโปรเจกต์ (Project Settings)** ติ๊กถูกที่ **"แสดงไฟล์ Manifest 'appsscript.json' ในตัวแก้ไข"**
3. กลับมาที่หน้า **ตัวแก้ไข (Editor)**:
   - นำโค้ดใน `appsscript.json` ไปวางทับไฟล์เดิม
   - นำโค้ดใน `Code.gs` ไปวางในไฟล์ `Code.gs`
   - กดปุ่ม `+` เลือก **HTML** ตั้งชื่อไฟล์ว่า `Index` แล้วนำโค้ดใน `Index.html` ไปวาง
4. กด **Save (บันทึกโปรเจกต์)**
5. กดปุ่ม **Deploy (การทำให้ใช้งานได้)** -> **New deployment** -> เลือกประเภท **Web app**
   - **Execute as**: Me
   - **Who has access**: Anyone (หรือ Only myself)
   - กด **Deploy** แล้วนำลิงก์ Web App ไปเปิดใช้งานได้ทันที
