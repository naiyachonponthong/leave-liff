# ระบบลาออนไลน์ผ่าน LINE

Google Apps Script + Google Sheets + LINE Messaging API + LIFF · Bootstrap 5 + jQuery + Select2

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
| `js_employee.html` | พนักงาน + ผูก LINE + โควตา |
| `js_leave.html` | ใบลา — สร้าง (Select2) / ค้นหา / รายละเอียด |
| `js_approval.html` | คิวอนุมัติ 2 ระดับ + Bulk approve/reject |
| `js_report.html` | รายงาน + Excel/PDF |
| `js_settings.html` | ตั้งค่า + อัพโหลดโลโก้ + ยกยอดวันลา + Rich Menu |

### GitHub Pages (LIFF)
| ไฟล์ | หน้าที่ |
|---|---|
| `liff.html` → `index.html` | หน้าพนักงานบน LINE LIFF (host บน GitHub Pages) |

> LIFF ต้อง host บน static domain (ไม่ใช่ GAS iframe) เพราะ LINE OAuth callback ทำงานไม่ได้ใน iframe

---

## ฟีเจอร์หลัก

### Admin / HR
- Dashboard: KPI 4 ตัว, Donut chart, Bar chart รายเดือน, **ปฏิทินการลาสัปดาห์นี้**
- คิวอนุมัติ: เลือก checkbox หลายรายการ → **Bulk อนุมัติ/ปฏิเสธ** พร้อมกัน
- ใบลา: **Select2** ค้นหาพนักงานได้ในช่อง dropdown
- ประเภทลา: กำหนด **Seniority Tiers** (โควตาตามอายุงาน)
- ยกยอดวันลาข้ามปี (carry forward) ด้วยปุ่มเดียว
- สร้าง **LINE Rich Menu** อัตโนมัติ
- หน้า **คู่มือการใช้งาน** แบบ tab (6 หัวข้อ)

### LINE LIFF (พนักงาน)
- ยื่นใบลา, ดูประวัติ, วันคงเหลือ, **ปฏิทินวันหยุด** (4 แท็บ)
- **กรองประวัติ** ตามสถานะด้วย chip filter
- **กดดูรายละเอียด** — bottom sheet เลื่อนขึ้น (เหตุผล, ผู้อนุมัติ, ความเห็น, เอกสาร)
- **ยกเลิกใบลา** ที่ยังรออนุมัติได้โดยตรงจาก LIFF
- Progress bar แสดง % วันลาที่ใช้ไปในแต่ละประเภท

### LINE Approval (ผู้อนุมัติ)
- เมื่อพนักงานยื่นลา ระบบส่ง **Flex Message** พร้อมปุ่ม **อนุมัติ / ไม่อนุมัติ** ให้ผู้อนุมัติทาง LINE
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

---

## โครงสร้างข้อมูล (Google Sheets — JSON-per-row)

| ชีต | ข้อมูล |
|---|---|
| Config | การตั้งค่าระบบ (1 แถว) |
| Users | บัญชีผู้ใช้ admin/hr/supervisor + session |
| Employees | พนักงาน + LINE userId + แผนก + role |
| Departments | แผนก + หัวหน้า |
| LeaveTypes | ประเภทการลา + โควตา + seniority_tiers |
| LeaveBalances | โควตาพนักงานรายปี (entitled/used/pending/remaining) |
| LeaveRequests | ใบลา + approvals array (JSON) |
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
       └─ พนักงานยกเลิก  ──►  CANCELLED
```

- **หัก pending** ตอนยื่น
- **ย้าย pending → used** ตอนอนุมัติสุดท้าย
- **คืน pending** ตอนปฏิเสธหรือยกเลิก
- แจ้งเตือน LINE ทุก transition (Flex Message + ปุ่ม postback)

---

## บทบาท (Role)

| Role | สิทธิ์ |
|---|---|
| `admin` | ทุกเมนู รวม ตั้งค่า |
| `hr` | dashboard, อนุมัติ, ใบลา, พนักงาน, ข้อมูลหลัก, รายงาน |
| `supervisor` | dashboard, อนุมัติ, ใบลา |
| `employee` | ใช้งานผ่าน LIFF เท่านั้น |

---

## หมายเหตุ

- ใช้กฎ escape `\x3c/` ในไฟล์ JS ทุกไฟล์ (ห้ามเปลี่ยนกลับเป็น `</` ใน string literal)
- ปรับเป็นอนุมัติ 1 ระดับได้ในหน้า ตั้งค่า
- LIFF API ใช้ `fetch` + `text/plain` body เพื่อหลีกเลี่ยง CORS preflight
- ฟังก์ชันที่เรียกได้จาก LIFF ต้องอยู่ใน `LIFF_API_WHITELIST` ใน Code.gs
