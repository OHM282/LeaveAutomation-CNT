# Release Notes: CNT STAFF TIME AUTOMATION

## เวอร์ชัน 1.2.8 (Universal Edition — Apple Liquid Glass Hybrid)
**วันที่ปล่อยอัปเดต:** 4 พฤศจิกายน 2026  
**ชื่อไฟล์เผยแพร่:** `LeaveAutomation.exe` | `LeaveAutomation.zip`  
**GitHub Repository:** `OHM282/LeaveAutomation-CNT`  
**สถานะการทดสอบ:** ผ่านการทดสอบระดับระบบ (Regression Test) 100% ครบทุกโมดูล

---

### 🌟 ไฮไลต์สำคัญในเวอร์ชัน 1.2.8

#### 1. เปลี่ยนชื่อเป็น Universal Edition อย่างเป็นทางการ
- ตัดรหัสโครงการ `1856` ออกจากชื่อไฟล์และระบบทั้งหมด เพื่อรองรับการนำไปใช้งานในทุกโครงการก่อสร้างและทุกส่วนงานของ **Christiani & Nielsen (Thai) PCL.** ทั่วประเทศ
- เปลี่ยนชื่อโปรแกรมหลักเป็น **`LeaveAutomation.exe`** และ **`LeaveAutomation.zip`**
- อัปเดตสเปกการบิลด์เป็น `LeaveAutomation.spec` และปรับแก้ระบบ Auto-Updater ตลอดจนตัวอย่างไฟล์ในหน้าต่างวิธีใช้ (Manual)

#### 2. ดีไซน์ใหม่: Apple Liquid Glass Hybrid Design (Phase 2 Refinement)
- **Windows DWM Dark Title Bar:** ผสานรอยต่อระหว่าง Title Bar ของ Windows 11 กับโทนสี Deep Navy (`#0a192f`) ด้วย Windows Desktop Window Manager API ทำให้หน้าต่างดูหรูหรา ไร้รอยต่อแบบ macOS Application
- **Workflow Stepper 6 ขั้นตอน:** แสดงลำดับขั้นตอนการทำงานอย่างเป็นระบบ (Source File ➔ Import Data ➔ Map Columns ➔ Save Mode ➔ Pre-Check ➔ Processing) พร้อมไฟสถานะแบบเรียลไทม์
- **Smart File Chips & Drop Zone:** เพิ่มแถบแสดงสถานะไฟล์ที่เลือก พร้อมข้อมูลนามสกุล ขนาดไฟล์ และรองรับการลากวาง (Drag & Drop) สะดวก รวดเร็ว
- **Developer-Grade Activity Log Console:** ปรับปรุงคอนโซลแสดงผลการทำงานให้มีมิติ Dark Theme สไตล์คอนโซลนักพัฒนา แยกสี Log Level (`[INFO]`, `[SUCCESS]`, `[WARN]`, `[ERROR]`) ชัดเจน พร้อมปุ่มคัดลอก บันทึกไฟล์ และล้างข้อความ
- **แก้ปัญหาปุ่มสเกลตัวอักษรโดนตัดขอบ:** ปรับความกว้างของ Segmented Button สำหรับผู้สูงวัยเป็น 185px แก้ปัญหาตัวเลข `100%`, `115%`, `125%` ขาดหายเป็น `.00%` ให้แสดงผลได้สมบูรณ์ คมชัด ทุกความละเอียดหน้าจอ
- **High Readability First:** รักษาความคมชัดของข้อความและคอนทราสต์ในทุกช่อง Input และตารางข้อมูล อ่านง่าย สบายตา สอดคล้องกับมาตรฐาน Apple HIG

#### 3. ปรับปรุง Logic คำนวณเวลาและการลา (Core Engine)
- **Auto-Fix เวลาเลิกงานวันหยุด/วันอาทิตย์ (15:00 น.):** แก้ปัญหากรณีพนักงานสแกนออกรอบบ่ายช่วง 12:00–16:00 น. ในวันที่มีเครื่องหมาย "น" หรือวันอาทิตย์ ระบบจะลงเวลา Check-Out ให้อัตโนมัติ ป้องกันไม่ให้ขึ้นข้อความ "ไม่ได้แสกน" และไม่ถูกหักเวลา 4 ชั่วโมงโดยไม่จำเป็น
- **Universal Column Detection:** ตรวจจับคอลัมน์อัตโนมัติรองรับแบบฟอร์มสแกนเวลาทั้งแบบเดิมและแบบใหม่ รวมถึงรองรับรูปแบบปี พ.ศ. และ ค.ศ.
- **ระบบตรวจสอบความผิดปกติ (Pre-Check 100% Preview):** จำลองผลการอ่านชีทและจับคู่พนักงานก่อนบันทึกจริง ป้องกันข้อมูลตกหล่น

#### 4. เสถียรภาพและการติดตั้ง (Zero-Dependency Standalone)
- **ฝังชุด Microsoft Visual C++ Runtime & UCRT DLLs ครบถ้วน:** ปิดโอกาสเกิดข้อผิดพลาด `Failed to load Python DLL python314.dll / LoadLibrary: The specified module could not be found` เมื่อเปิดบนเครื่องลูกข่ายที่ไม่มี Python หรือ VC++ Redistributable
- **ปิดระบบบีบอัด UPX:** แก้ปัญหา Antivirus บน Windows แจ้งเตือน False Positive และเพิ่มความเร็วในการเปิดโปรแกรม (Startup Time)

---

### 📦 ไฟล์สำหรับติดตั้งและอัปเดต
| ไฟล์ | รายละเอียด | ขนาด |
| :--- | :--- | :--- |
| **`LeaveAutomation.exe`** | ไฟล์โปรแกรมหลักสำหรับเปิดใช้งานบน Windows (Standalone) | ~127 MB |
| **`LeaveAutomation.zip`** | ไฟล์บีบอัดสำหรับดาวน์โหลดหรือส่งต่อผ่านระบบเครือข่าย/อีเมล | ~125 MB |

---

### 🛠️ วิธีการอัปเดต
1. **ผ่านโปรแกรม:** เปิดโปรแกรมเดิม แล้วคลิกปุ่ม **"อัปเดต v1.2.8"** ที่มุมขวาบน ระบบจะดาวน์โหลดและติดตั้งให้อัตโนมัติ
2. **แบบ Manual:** ก๊อปปี้ไฟล์ `LeaveAutomation.exe` หรือแตกไฟล์จาก `LeaveAutomation.zip` ไปวางทับโฟลเดอร์เดิม แล้วดับเบิลคลิกใช้งานได้ทันที
