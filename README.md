# ระบบลาออนไลน์ผ่าน LINE

Google Apps Script + Google Sheets + LINE Messaging API + LIFF · Bootstrap 5 + jQuery + Select2

---

## ไฟล์โปรเจกต์

### Google Apps Script (GAS)

| ไฟล์ | หน้าที่ |
|---|---|
| `Code.gs` | Backend ทั้งหมด — API, webhook, session, notification |
| `index.html` | Shell หน้าแอดมิน (login + layout + modals) |
| `css.html` | ธีม OKLCH + dark mode + skeleton + Select2 |
| `js_core.html` | Auth / session / router / theme / helpers |
| `js_dashboard.html` | KPI + Chart.js donut/bar + ปฏิทินสัปดาห์ |
| `js_master.html` | ประเภทลา (+ seniority tiers) / แผนก / วันหยุด |
| `js_employee.html` | พนักงาน + ผูก LINE + โควตา + Import CSV |
| `js_leave.html` | ใบลา — สร้าง (Select2) / ค้นหา / รายละเอียด |
| `js_approval.html` | คิวอนุมัติ 2 ระดับ + Bulk approve/reject |
| `js_report.html` | รายงาน + Excel/PDF |
| `js_settings.html` | ตั้งค่า + อัพโหลดโลโก้ + ยกยอดวันลา + Rich Menu |
| `js_users.html` | จัดการบัญชีผู้ใช้ระบบ (admin only) |
| `js_calendar.html` | ปฏิทินการลารายเดือน (grid view) |
| `js_profile.html` | เปลี่ยนรหัสผ่านตนเอง |

### GitHub Pages (LIFF)

| ไฟล์ | หน้าที่ |
|---|---|
| `liff.html` → `index.html` | หน้าพนักงานบน LINE LIFF (host บน GitHub Pages) |

> LIFF ต้อง host บน static domain (ไม่ใช่ GAS iframe) เพราะ LINE OAuth callback ทำงานไม่ได้ใน iframe

---

## ฟีเจอร์หลัก

### Admin / HR
- **Dashboard:** KPI 4 ตัว, Donut chart, Bar chart รายเดือน, ปฏิทินการลาสัปดาห์นี้
- **ปฏิทินการลา:** grid view รายเดือน เลื่อนเดือนไปหน้า/ย้อนหลัง แสดงชื่อผู้ลาในแต่ละวัน + legend สี + วันหยุด
- **คิวอนุมัติ:** เลือก checkbox หลายรายการ → Bulk อนุมัติ/ปฏิเสธพร้อมกัน
- **ใบลา:** Select2 ค้นหาพนักงานได้ในช่อง dropdown
- **พนักงาน:** เพิ่ม/แก้ไข/ลบ ผูก LINE + จัดการโควตา + **Import จาก CSV** พร้อม preview ก่อนนำเข้า
- **ประเภทลา:** กำหนด Seniority Tiers (โควตาตามอายุงาน)
- **แผนก:** กำหนดหัวหน้าได้ **หลายคน** (multi-select) — ทุกคนรับ Flex Message อนุมัติ
- **จัดการผู้ใช้ (admin only):** CRUD บัญชี admin/hr/supervisor — ชื่อ, username, รหัสผ่าน, อีเมล, สถานะ
- **เปลี่ยนรหัสผ่าน:** ทุก role เปลี่ยนรหัสผ่านตัวเองได้จากเมนู user dropdown
- **ยกยอดวันลาข้ามปี** (carry forward) ด้วยปุ่มเดียว
- **แจ้งเตือนอัตโนมัติ:** ระบบแจ้งเตือน LINE ทุก 09:00 น. เมื่อมีใบลาค้างอนุมัติ > 2 วัน
- สร้าง LINE Rich Menu อัตโนมัติ

### LINE LIFF (พนักงาน)
- **ยื่นใบลา** พร้อมแนบไฟล์เอกสาร
- **ประวัติการลา** กรองตามสถานะด้วย chip filter
- **วันลาคงเหลือ** พร้อม progress bar แสดง % ที่ใช้ไป
- **สถิติรายบุคคล** แยกตามประเภทลา พร้อม progress bar และ % — เลือกดูย้อนหลังได้
- **ปฏิทินวันหยุด** รายปี
- **ดูรายละเอียดใบลา** — bottom sheet เลื่อนขึ้น (เหตุผล, ผู้อนุมัติ, ความเห็น, เอกสาร)
- **ยกเลิกใบลา** ที่ยังรออนุมัติ พร้อมระบุ **เหตุผลการยกเลิก**

### LINE Approval (ผู้อนุมัติ)
- เมื่อพนักงานยื่นลา ระบบส่ง **Flex Message** พร้อมปุ่ม **อนุมัติ / ไม่อนุมัติ** ให้ผู้อนุมัติทาง LINE
- หัวหน้าแผนกหลายคนรับ Flex Message พร้อมกันได้
- กดปุ่มได้เลยในแชท — ไม่ต้องเปิดหน้าเว็บ
- ระบบตอบกลับผลในแชทและ push แจ้งพนักงานทันที

---

## ติดตั้ง

### 1. GAS Backend

```
1. สร้าง Google Sheet → Extensions → Apps Script
2. วาง Code.gs และสร้างไฟล์ HTML ทุกชื่อตามตาราง (ไม่ต้องใส่ .html ในชื่อฝั่ง GAS)
3. รันฟังก์ชัน setup() หนึ่งครั้ง — สร้าง 11 ชีต + seed ข้อมูลตัวอย่าง
4. Deploy → New deployment → Web app
   Execute as: Me · Access: Anyone
5. คัดลอก Web App URL
```

บัญชีเริ่มต้น: `admin / admin123` · `hr / hr123` · `supervisor / sup123`

### 2. GitHub Pages (LIFF)

```
1. สร้าง GitHub repo ใหม่ (public)
2. วาง liff.html → เปลี่ยนชื่อเป็น index.html
3. Settings → Pages → Deploy from main branch
4. URL: https://{username}.github.io/{repo}/
```

### 3. LINE Developers Console

```
Messaging API Channel:
  - Webhook URL = {GAS Web App URL}
  - เปิด "Use webhook"
  - ปิด Auto-reply และ Greeting message

LIFF:
  - Add LIFF app → Size: Full
  - Endpoint URL = https://{username}.github.io/{repo}/
  - Scope: profile
  - คัดลอก LIFF ID
```

### 4. ตั้งค่าในระบบ

```
1. เข้าหน้าแอดมิน → เมนู ตั้งค่า
2. ใส่ Channel Access Token, Channel Secret, LIFF ID
3. กด บันทึกการตั้งค่า
4. กด ตั้งค่า Rich Menu (สร้างเมนู LINE อัตโนมัติ)
```

### 5. ตั้งค่าแจ้งเตือนอัตโนมัติ (ทำครั้งเดียว)

```
1. เปิด Apps Script Editor
2. เลือก Run → setupRemindTrigger()
3. ระบบจะสร้าง Time-driven trigger แจ้งเตือนทุกวัน 09:00 น.
   เมื่อมีใบลาค้างอนุมัติเกิน 2 วัน
```

---

## Import พนักงานจาก CSV

### รูปแบบไฟล์ CSV

```csv
คำนำหน้า,ชื่อ,นามสกุล,ชื่อเล่น,รหัสพนักงาน,แผนก,ตำแหน่ง,ประเภท,เบอร์โทร,อีเมล,วันเริ่มงาน
นาย,สมชาย,ใจดี,ชาย,EMP001,ฝ่ายขาย,พนักงานขาย,ประจำ,081-000-0001,somchai@example.com,2024-01-01
นางสาว,สมหญิง,รักงาน,หญิง,EMP002,ฝ่ายบัญชี,นักบัญชี,ประจำ,081-000-0002,somying@example.com,2024-02-01
```

**คอลัมน์บังคับ:** `ชื่อ`, `นามสกุล`  
**คอลัมน์ไม่บังคับ:** ที่เหลือทั้งหมด (ถ้าไม่ใส่จะเว้นว่างไว้)  
ชื่อแผนกต้องตรงกับที่มีในระบบพอดี เพื่อให้ผูก department_id ได้ถูกต้อง

---

## โครงสร้างข้อมูล (Google Sheets — JSON-per-row)

| ชีต | ข้อมูลสำคัญ |
|---|---|
| Config | การตั้งค่าระบบ (1 แถว) |
| Users | บัญชีผู้ใช้ admin/hr/supervisor — username, password, role, line_user_id |
| Employees | พนักงาน + LINE userId + department_id |
| Departments | แผนก + `head_user_ids` (array ของ Users.id) |
| LeaveTypes | ประเภทการลา + โควตา + seniority_tiers |
| LeaveBalances | โควตาพนักงานรายปี (entitled/used/pending/remaining) |
| LeaveRequests | ใบลา + approvals array (JSON) + cancel_reason |
| Holidays | วันหยุดราชการ/บริษัทรายปี |
| Sessions | session token (TTL 8 ชม.) |
| Errors | log ข้อผิดพลาด |
| Notifications | log การแจ้งเตือน LINE |

---

## Flow การลา

```
ยื่น (เว็บ Admin / LINE LIFF)
  └─ SUBMITTED
       ├─ หัวหน้าอนุมัติ (เว็บ หรือ กดปุ่มใน LINE)
       │    └─ L1_APPROVED  ──►  HR อนุมัติ  ──►  APPROVED
       ├─ ปฏิเสธ  ──►  REJECTED
       └─ พนักงานยกเลิก (+ เหตุผล)  ──►  CANCELLED
```

- **หัก pending** ตอนยื่น
- **ย้าย pending → used** ตอนอนุมัติสุดท้าย
- **คืน pending** ตอนปฏิเสธหรือยกเลิก
- แจ้งเตือน LINE ทุก transition (Flex Message + ปุ่ม postback)

---

## บทบาท (Role)

| Role | เมนูที่เข้าถึงได้ |
|---|---|
| `admin` | ทุกเมนู รวม จัดการผู้ใช้ และ ตั้งค่า |
| `hr` | dashboard, ปฏิทิน, อนุมัติ, ใบลา, พนักงาน, ข้อมูลหลัก, รายงาน |
| `supervisor` | dashboard, ปฏิทิน, อนุมัติ, ใบลา (เฉพาะแผนกตนเอง) |
| employee | ใช้งานผ่าน LINE LIFF เท่านั้น |

### การแยกข้อมูล Supervisor
- Supervisor เห็นเฉพาะใบลาของแผนกที่ตนเป็นหัวหน้า
- Dashboard, คิวอนุมัติ, รายการใบลา — กรองตาม `head_user_ids` โดยอัตโนมัติ
- หัวหน้าแผนกต้องเป็น Users (role = supervisor) ไม่ใช่ Employees

---

## API สำคัญ (Code.gs)

| ฟังก์ชัน | สิทธิ์ | หน้าที่ |
|---|---|---|
| `setup()` | - | สร้างชีตและ seed ข้อมูล (รันครั้งเดียว) |
| `login(username, password)` | - | เข้าสู่ระบบ |
| `changePassword(token, currentPw, newPw)` | ทุก role | เปลี่ยนรหัสผ่านตัวเอง |
| `getUsers(token)` | admin | ดูบัญชีผู้ใช้ทั้งหมด |
| `saveUser(token, payload)` | admin | เพิ่ม/แก้ไขบัญชี |
| `deleteUser(token, id)` | admin | ลบบัญชี |
| `importEmployeesCSV(token, rows)` | admin/hr | นำเข้าพนักงานจาก CSV |
| `getMonthlyCalendar(token, year, month)` | ทุก role | ใบลา APPROVED รายเดือน |
| `remindPendingApprovals()` | trigger | แจ้งเตือนใบลาค้างอนุมัติ |
| `setupRemindTrigger()` | - | สร้าง time-driven trigger (รันครั้งเดียว) |
| `liffGetMyStats(lineUserId, year)` | LIFF | สถิติการลารายบุคคล |
| `liffCancelLeave(lineUserId, leaveId, reason)` | LIFF | ยกเลิกใบลาพร้อมเหตุผล |

---

## หมายเหตุสำหรับนักพัฒนา

- ใช้กฎ escape `\x3c/` ในไฟล์ JS ทุกไฟล์ (ห้ามเปลี่ยนกลับเป็น `</` ใน string literal ภายใน GAS HTML)
- ปรับเป็นอนุมัติ 1 ระดับได้ในหน้า ตั้งค่า
- LIFF API ใช้ `fetch` + `text/plain` body เพื่อหลีกเลี่ยง CORS preflight
- ฟังก์ชันที่เรียกได้จาก LIFF ต้องอยู่ใน `LIFF_API_WHITELIST` ใน Code.gs
- `head_user_ids` เป็น array ใน Departments — backward compat กับ `head_user_id` (string เดี่ยว) ยังทำงานได้
- Permissions ของ user ถูก derive จาก `CONFIG.USER_ROLES` ตอน login ทุกครั้ง ไม่ได้อ่านจาก Sheets
