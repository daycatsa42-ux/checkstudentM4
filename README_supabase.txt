# ระบบเช็คชื่อ ม.4 - Supabase Version

## วิธีใช้
1. เปิดไฟล์ `index.html`
2. ใส่ค่า Supabase ตรงบรรทัด:
   - SUPABASE_URL
   - SUPABASE_ANON_KEY
3. อัปโหลด `index.html` ไป GitHub Pages

## ตาราง Supabase ที่ต้องมี
students:
- id
- room
- no
- name

attendance:
- id
- date
- room
- student_no
- student_name
- status
- timestamp

หมายเหตุ: ต้องมี unique index: date, room, student_no
