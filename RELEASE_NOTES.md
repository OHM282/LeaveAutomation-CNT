# Release Notes: CNT STAFF TIME AUTOMATION

## เวอร์ชัน 1.2.8 (Universal Edition — Apple Liquid Glass Soft Light Premium Edition)
**วันที่ปล่อยอัปเดต:** 5 พฤศจิกายน 2026  
**ชื่อไฟล์เผยแพร่:** `LeaveAutomation.exe` | `LeaveAutomation.zip`  
**GitHub Repository:** `OHM282/LeaveAutomation-CNT`  
**สถานะการทดสอบ:** ผ่านการทดสอบระดับระบบ (Regression Test) 100% ครบทุกโมดูล

---

### 🌟 ไฮไลต์สำคัญในเวอร์ชัน 1.2.8 (Soft Light Visual Correction)

#### 1. ยกระดับ Visual Language เป็น Apple Liquid Glass Soft Light
- **Soft Neutral Canvas (`#F5F6FA`):** เปลี่ยนพื้นหลังจากมืดทึบเป็นโทนสว่างนุ่มนวล สบายตา ไม่อึดอัด
- **Light Frosted Header (`#FFFFFF` + Subdivider `#E5E7EB`):** แถบหัวเรื่องสไตล์ macOS Light Frosted Glass พร้อมฟอนต์เข้ม คอนทราสต์ชัดเจน และโลโก้ CN ไร้กรอบการ์ดขาว
- **Elevated Floating White Cards (`#FFFFFF` + Border `#D8DCE3`):** การ์ดพื้นขาวลอยตัวพร้อมขอบบางโปร่งแสง ให้มิติและความลึก ไม่แบนเรียบและไร้ขอบสีฟ้าเข้มหนาเตอะ
- **Graphite Developer Console (`#1D1D1F`):** คอนโซลสไตล์ macOS Developer Terminal พื้นกราไฟต์เข้มระดับพรีเมียม (ไม่ใช่ดำสนิท `#000000`) ข้อความสีเทานุ่มนวล (`#E5E5EA`) พร้อมระบบจำแนกสีสถานะ (Success `#30A46C`, Warning `#D99A24`, Error `#E5484D`, Info `#007AFF`)
- **Minimal Segmented Stepper:** ลำดับขั้นตอน 6 ขั้นตอน สไตล์ Segmented Glass พร้อมสีสถานะชัดเจน (เขียว Soft Emerald เมื่อเสร็จ, ฟ้า Apple Blue เมื่อพร้อม, ส้ม Amber เมื่อกำลังทำ, แดง Soft Red เมื่อขัดข้อง)
- **Windows 11 Native DWM Light Titlebar:** ซิงค์ไตเติลบาร์ของ Windows 11 ให้เป็นสีขาวนวล `#FFFFFF` เข้ากับตัวโปรแกรม 100% ตัดแถบสีแดงหรือสีที่แปลกแยกออกอย่างสิ้นเชิง

#### 2. แก้ไข Bug & Typography Hierarchy
- **แก้ปัญหาปุ่มสเกลตัวอักษรโดนตัดขอบ (`.00%` -> `100%`):** ขยายความกว้าง Segmented Button เป็น 220px และปรับ `dynamic_resizing=False` ป้องกันระบบ DPI Scaling บน Windows ตัดตัวเลขตัวหน้าทิ้ง พร้อมระบบ Active Pill Highlight สีขาวคมชัด
- **ปุ่ม Action Buttons มิติ Apple:** ปุ่ม "Pre-Check Preview" สไตล์ Glass Secondary และปุ่ม "▶ เริ่มประมวลผลทันที" ในสี Soft Emerald (`#30A46C`) พร้อมสถานะ Loading และระบบป้องกันการกดซ้ำระหว่างประมวลผล
- **Smart File Hub Drop Zone:** แสดงสถานะไฟล์ ชื่อไฟล์ ขนาดไฟล์ (KB/MB) และปุ่ม Browse / Change แยกออกจากพื้นที่ข้อมูลอย่างชัดเจน

#### 3. ปรับปรุง Logic คำนวณเวลาและการลา (Core Engine 100% Preserved)
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
