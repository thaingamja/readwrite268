#  ระบบบันทึกและประเมินผลการอ่าน การเขียน การฟัง และการพูด (Student Evaluation System)

![Version](https://img.shields.io/badge/version-5.1.0-blue.svg) ![Status](https://img.shields.io/badge/status-stable-green.svg) ![Platform](https://img.shields.io/badge/platform-Google%20Apps%20Script-orange.svg)

Web Application สำหรับบันทึกคะแนนและประเมินผลนักเรียนรายบุคคล รองรับการทำงานพร้อมกันหลายคน (Concurrency) มีระบบ Admin, การแปลผลอัตโนมัติ (Auto Interpretation), และการออกรายงาน (PDF/Excel)

---

##  ฟีเจอร์หลัก (Features)

*  **ระบบยืนยันตัวตน (Authentication):** แยกสิทธิ์การใช้งานระหว่าง ครูประจำวิชา และ Admin
*  **แปลผลอัตโนมัติ (Real-time Interpretation):** แสดงผลประเมิน (ดีมาก, ดี, พอใช้) ทันทีที่กรอกคะแนน
*  **High Performance:** รองรับการใช้งานพร้อมกันจำนวนมากด้วยเทคนิค Batch Update และ LockService
*  **Responsive UI:** ใช้งานได้ดีทั้งบนคอมพิวเตอร์และแท็บเล็ต
*  **Dashboard & Reports:** สรุปผลแบบกราฟวงกลม และส่งออกรายงานเป็น Excel/PDF ได้ทันที
*  **Dynamic Configuration:** ปรับเปลี่ยนเกณฑ์คะแนนและหัวข้อประเมินได้เองผ่าน Google Sheets โดยไม่ต้องแก้โค้ด

---

##  การติดตั้งและเตรียมฐานข้อมูล (Database Setup)

ระบบนี้ใช้ **Google Sheets** เป็นฐานข้อมูลหลัก กรุณาสร้างไฟล์ใหม่และสร้างแผ่นงาน (Tabs) ดังต่อไปนี้:

### 1. ชีต `Users` (ผู้ใช้งาน)
สำหรับเก็บข้อมูลล็อกอินของครูและผู้ดูแลระบบ

| Column A | Column B | Column C | Column D | Column E |
| :--- | :--- | :--- | :--- | :--- |
| **Username** | **Password** | **Name** | **Room** | **Level** |
| `kru_som` | `1234` | ครูสมศรี | ป.1/1 | ป.1 |
| `admin` | `1234` | Admin System | All | **Admin** |

> **Note:** หากต้องการให้เป็น Admin ให้ใส่ช่อง Level ว่า `Admin`

### 2. ชีต `Settings` (การตั้งค่าหัวข้อประเมิน)
กำหนดหัวข้อที่จะให้กรอกคะแนน (สามารถเพิ่มลดได้ตามใจชอบ)

| Column A | Column B | Column C | Column D | Column E | Column F |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Level** | **Header Name** | **Max Score** | **Data Col Index** | **Group** | **Criteria ID** |
| ป.1 | อ่านออกเสียง | 10 | 7 | การอ่าน | C_STD_10 |
| ป.1 | เขียนคำ | 20 | 8 | การเขียน | C_STD_20 |
| ป.1 | **รวม** | 30 | 9 | ผลรวม | C_SUM_30 |

> **Tips:**
> * **Data Col Index:** คือเลขคอลัมน์ในชีต Data ที่จะให้บันทึก (A=1, B=2, ..., G=7)
> * **Auto Sum:** หากตั้งชื่อ Header Name ว่า **"รวม"** ระบบจะคำนวณผลรวมให้อัตโนมัติและห้ามแก้ไขช่องนั้น

### 3. ชีต `Criteria` (เกณฑ์การประเมิน)
กำหนดช่วงคะแนนสำหรับการแปลผล

| Column A | Column B | Column C | Column D |
| :--- | :--- | :--- | :--- |
| **Criteria ID** | **Min Score** | **Max Score** | **Label** |
| C_STD_10 | 8 | 10 | ดีมาก |
| C_STD_10 | 5 | 7 | ดี |
| C_STD_10 | 0 | 4 | ปรับปรุง |

### 4. ชีต `Data` (ข้อมูลนักเรียนต้นฉบับ)
ข้อมูลรายชื่อนักเรียนทั้งหมด (Master Data) ระบบจะดึงจากที่นี่ไปสร้างชีตแยกรายห้องให้อัตโนมัติ

| Col A | Col B | Col C | Col D | Col E | Col F | Col G... |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **ID** | **Title** | **Name** | **Surname** | **Room** | **Teacher** | *(Score Cols)* |
| 1001 | ด.ช. | สมชาย | ใจดี | ป.1/1 | ครูสมศรี | *(เว้นว่าง)* |
| 1002 | ด.ญ. | มานี | รักเรียน | ป.1/2 | ครูมานะ | *(เว้นว่าง)* |

---

##  การ Deploy (Deployment)

1.  เปิด Google Sheets ไปที่ `Extensions` > `Apps Script`
2.  คัดลอกโค้ดจากไฟล์ `Code.gs` ลงไป
3.  กด **Deploy** > **New Deployment**
4.  เลือก type เป็น **Web App**
5.  ตั้งค่า:
    * **Execute as:** `Me` (ตัวคุณเอง)
    * **Who has access:** `Anyone` (ทุกคน) *สำคัญมากเพื่อเลี่ยง CORS*
6.  คัดลอก **Web App URL** ที่ได้ ไปใส่ในตัวแปร `SCRIPT_URL` ในไฟล์ `Index.html`

---

##  เทคโนโลยีที่ใช้ (Tech Stack)

* **Frontend:** HTML5, Tailwind CSS, JavaScript (ES6+)
* **Backend:** Google Apps Script (GAS)
* **Database:** Google Sheets
* **Libraries:**
    * `SweetAlert2` (Notifications)
    * `Chart.js` (Dashboard)
    * `SheetJS (XLSX)` (Excel Export/Import)
    * `PDFMake` (PDF Generation with Thai Fonts)

---

##  ผู้พัฒนา (Developer)

พัฒนาโดย **waroonporn rattanabhutchai สพป.สบ.2**
Full Stack Developer

&copy; 2026 All Rights Reserved.
