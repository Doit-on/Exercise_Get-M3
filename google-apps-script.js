/**
 * =========================================================================
 * Google Apps Script for Get INVOLVED! M.3 Online Exercise & E-Certificate System
 * ระบบแบบฝึกหัดออนไลน์และออกเกียรติบัตรอิเล็กทรอนิกส์ Get INVOLVED! ม.3 (ทวพ. - TWP)
 * เวอร์ชัน 2.0 (Account Authentication & Automated E-Certificate)
 * =========================================================================
 * 
 * คุณสมบัติสำคัญของระบบ:
 * 1. ฐานข้อมูลนักเรียน (Students_DB): อ้างอิงด้วยอีเมล (Email) และรหัสผ่าน ป้องกันคะแนนสับสน/ไม่ต่อเนื่อง
 *    - สามารถคัดลอกรายชื่อจาก Excel (.xlsx) มาวางในชีต Students_DB เพื่อสร้างฐานข้อมูลการ Login ได้ทันที
 * 2. บันทึกผลคะแนนรายข้อ (Scores_Log): เก็บประวัติการส่งงาน คะแนน วันเวลา และคำแนะนำครู
 * 3. ระบบออกเกียรติบัตรอัตโนมัติ (Certificates): เมื่อนักเรียนทำแบบฝึกหัดครบทั้ง 4 Units (5, 6, 7, 8)
 *    - บันทึกรหัสเกียรติบัตร (Serial Number), คะแนนเฉลี่ยรวม, วันที่ออกเกียรติบัตร
 *    - มีระบบส่งอีเมลแจ้งเตือนเกียรติบัตรถึงนักเรียนอัตโนมัติ (Automated Email Dispatch)
 * 4. การรักษาความปลอดภัยครู (Teacher_Config): เก็บรหัสผ่านครูในระบบ ปิดซ่อนรหัสผ่านจากหน้าจอ UI
 * 
 * วิธีการติดตั้ง (3 ขั้นตอนง่ายๆ):
 * 1. เปิด Google Sheets ใหม่ (ตั้งชื่อว่า TWP_GetInvolved_M3_Database) -> ไปที่เมนู "ส่วนขยาย" (Extensions) -> "Apps Script"
 * 2. ลบโค้ดเดิมออกทั้งหมด แล้ววางโค้ดไฟล์นี้ลงไปทั้งหมด -> กดบันทึก (Save)
 * 3. เลือกฟังก์ชัน "setupAllSheets" จากเมนูด้านบน แล้วกดปุ่ม "เรียกใช้" (Run) 1 ครั้ง
 *    (เพื่อสร้างและจัดรูปแบบชีตทั้ง 4 แผ่นให้อัตโนมัติ สวยงาม พร้อมใช้งานทันที)
 * 4. กดปุ่ม "ทำให้ใช้งานได้" (Deploy) -> "การทำให้ใช้งานได้รายการใหม่" (New deployment)
 *    - เลือกประเภท: "เว็บแอป" (Web app)
 *    - ดำเนินการในฐานะ: "ฉัน" (Execute as: Me)
 *    - ผู้ที่มีสิทธิ์เข้าถึง: "ทุกคน" (Who has access: Anyone) **สำคัญมาก ต้องเป็นทุกคน**
 *    - กด "ทำให้ใช้งานได้" (Deploy) แล้วคัดลอก URL ของเว็บแอป นำไปใส่ใน WebApp
 * =========================================================================
 */

// ชื่อชีตหลักในระบบ

/**
 * เพิ่มเมนูพิเศษบน Google Sheets อัตโนมัติเมื่อเปิดชีต
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎓 ระบบ TWP ม.3')
    .addItem('⚡ ตั้งค่าและสร้างชีตทั้งหมด (Setup Database)', 'setupAllSheets')
    .addToUi();
}

const SHEET_STUDENTS = "Students_DB";
const SHEET_NAME_BASE = "Name_base";
const SHEET_LOGIN_LOGS = "Login_Logs";
const SHEET_SCORES = "Scores_Log";
const SHEET_CERTS = "Certificates";
const SHEET_CONFIG = "Teacher_Config";

/**
 * ฟังก์ชันสร้างและจัดรูปแบบตารางทั้งหมดอัตโนมัติ (รันครั้งแรกเมื่อติดตั้ง)
 */
function setupAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // นำเข้ารายชื่อนักเรียนตั้งต้น 53 คนจากเอกสารโครงการ (PDF Roster)
  const initialStudents = [
  [
    "STD-001",
    "พุฒิภัทร สิงห์บุตรดี",
    "รร.เมืองกระบี่ (กระบี่)",
    "ม.3",
    "1",
    "sawaddeeja020856@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-002",
    "ด.ช.ปัณณ์ ปณกร",
    "โรงเรียนเมืองกระบี่ (กระบี่)",
    "ม.3",
    "2",
    "punnawit060257@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-003",
    "ด.ญ.ณัฐนิชา บุญปลอด",
    "โรงเรียนเมืองกระบี่ (กระบี่)",
    "ม.3",
    "3",
    "natnichabunplod470@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-004",
    "ด.ญ นวนิดา เจริญพร",
    "โรงเรียนเมือง (กระบี่)",
    "ม.3",
    "4",
    "tusdsannee3212@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-005",
    "ด.ญ.ปัณณภัสร์ เพชรแขก",
    "เมืองกระบี่ (กระบี่)",
    "ม.3",
    "5",
    "chanyanuch2527@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-006",
    "จิรัตน์กาล นาควิลัย",
    "โรงเรียนเมืองกระบี่ (กระบี่)",
    "ม.3",
    "6",
    "ciratntikalnakhwilay@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-007",
    "ด.ญ.กุลทณพัชร์ ทิพย์มงคล",
    "โรงเรียนเมืองกระบี่ (กระบี่)",
    "ม.3",
    "7",
    "jaonai1153@hotmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-008",
    "Phattarawadee Petchkerd",
    "Muangkrabi School (กระบี่)",
    "ม.3",
    "8",
    "Phattarawadee@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-009",
    "ด.ญ.ธนัชชา ถ่ายย้วนกุล",
    "เมืองกระบี่ (กระบี่)",
    "ม.3",
    "9",
    "thanutcha.tai6486@ndlp.go.th",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-010",
    "ภูริชชญา เนาวรัตน์",
    "เมืองกระบี่ (กระบี่)",
    "ม.3",
    "10",
    "Mikixdnoobztv@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-011",
    "ด.ช ธนภัทร คงสีลัง",
    "โรงเรียนเมืองกระบี่ (กระบี่)",
    "ม.3",
    "11",
    "tonnum650@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-012",
    "ด.ญ.ณพัชชนันท์ สิงห์อุไร",
    "โรงเรียนบดินทรเดชา (สิงห์ สิงหเสนี) 2 (กรุงเทพมหานคร)",
    "ม.3",
    "12",
    "naphatchanansingurai@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-013",
    "ด.ช.พุฒิพัชร์ เพ็ชรจำรัส",
    "โรงเรียนสาวิทยา (กรุงเทพมหานคร)",
    "ม.3",
    "13",
    "puttisa22@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-014",
    "ด.ช.จารุกิตติ์ แสนบุญรัตน์",
    "โรงเรียนปะกงพิทยาคม (พังงา)",
    "ม.3",
    "14",
    "phupha02022555@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-015",
    "จารุกิตติ์ แสนบุญรัตน์",
    "โรงเรียนปะกงพิทยาคม (พังงา)",
    "ม.3",
    "15",
    "phupha02022555",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-016",
    "ด.ช.พีรศุภ รักมิตร",
    "โรงเรียนเมืองกระบี่ (สระบุรี)",
    "ม.3",
    "16",
    "rphirsuph@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-017",
    "ด.ญ.ปุณณดา โกละกะ",
    "สุราษฎร์พิทยา (สุราษฎร์ธานี)",
    "ม.3",
    "17",
    "punndakolaka351@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-018",
    "ปุณณพัทธ ชูเพชร์",
    "สุราษฎร์พิทยา (สุราษฎร์ธานี)",
    "ม.3",
    "18",
    "Punn3101@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-019",
    "ด.ญ ปิยฐดา โปยขุนทด",
    "โรงเรียนโชคชัยสามัคคี (นครราชสีมา)",
    "ม.3",
    "19",
    "mamew1025@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-020",
    "ด.ญ.พลอยชมพู บุดดาแพน",
    "โรงเรียนโชคชัยสามัคคี (นครราชสีมา)",
    "ม.3",
    "20",
    "Phoychomphu10@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-021",
    "พิชญธิดา ครองยศ",
    "โกสุมวิทยาสรรค์ (มหาสารคาม)",
    "ม.3",
    "21",
    "Phichayathida1002@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-022",
    "พิชญดา อรุณรุ่งก้าวไกล",
    "โกสุมวิทยาสรรค์ (มหาสารคาม)",
    "ม.3",
    "22",
    "fahsai240855@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-023",
    "ธมนวรรณ ดวงกางใต้",
    "โรงเรียนโกสุมวิทยาสรรค์ (มหาสารคาม)",
    "ม.3",
    "23",
    "nxngmaprang56@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-024",
    "ด.ญสาวินี บุญเรืองขาว",
    "โรงเรียนเมืองกระบี่ (กระบี่)",
    "ม.3",
    "24",
    "sawinibuyreuxngkhaw705@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-025",
    "นางสาวอริสสรา อินทรสุวรรณ",
    "โรงเรียนสตรีทุ่งสง (นครศรีธรรมราช)",
    "ม.3",
    "25",
    "ab0930724152@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-026",
    "กัลยกร วัยวัฒน์",
    "โรงเรียนสตรีทุ่งสง (นครศรีธรรมราช)",
    "ม.3",
    "26",
    "p0951090517@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-027",
    "Chuthamat Phankaeo",
    "สตรีทุ่งสง (นครศรีธรรมราช)",
    "ม.3",
    "27",
    "chuthamatphankaeo@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-028",
    "จุฑามาศ พันธ์แก้ว",
    "สตรีทุ่งสง (นครศรีธรรมราช)",
    "ม.3",
    "28",
    "chuthamatphankaeo@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-029",
    "ด.ญ.พิมพ์ศญา สามเกษร",
    "ธัญบุรี (ปทุมธานี)",
    "ม.3",
    "29",
    "tbs41951@tbs.ac.th",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-030",
    "Chayanin Patchaphan",
    "โรงเรียนดีบุกพังงาวิทยายน (พังงา)",
    "ม.3",
    "30",
    "yaya.chayanin0109@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-031",
    "น.ส.ปัณฑิตา สิงห์อุไร",
    "นวมินทราชินูทิศสตรีวิทยา2 (กรุงเทพมหานคร)",
    "ม.4",
    "31",
    "pantitasingurai@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-032",
    "นางสาวบุญชญา ไชยแก่น",
    "โรงเรียนเชียงใหม่คริสเตียน (เชียงใหม่)",
    "ม.4",
    "32",
    "chiykaenbuychya@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-033",
    "นายดนุสรณ์ เจ็กมะดัน",
    "โรงเรียนเสิงสาง (นครราชสีมา)",
    "ม.4",
    "33",
    "earthbaby5454@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-034",
    "Natakon Phanchanghan",
    "โรงเรียนสตรีทุ่งสง (นครศรีธรรมราช)",
    "ม.4",
    "34",
    "nattakorn.pk393@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-035",
    "นางสาววราภรณ์ ธิโกศรี",
    "โรงเรียนธัญบุรี (ปทุมธานี)",
    "ม.4",
    "35",
    "tbs41338@tbs.ac.th",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-036",
    "สุกฤตา เลิกสันเทียะ",
    "โรงเรียนธัญบุรี (ปทุมธานี)",
    "ม.4",
    "36",
    "tbs41299@tbs.ac.th",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-037",
    "ศศิธร",
    "ธัญบุรี (ปทุมธานี)",
    "ม.4",
    "37",
    "tbs43594@tbs.ac.th",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-038",
    "กัญญ์ณัชชา แซ่ชิ้น",
    "โรงเรียนธัญบุรี (ปทุมธานี)",
    "ม.4",
    "38",
    "arisasarar@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-039",
    "นลพรรณ์ แสงพงษ์พิทยา",
    "บดินทรเดชา (สิงห์ สิงหเสนี) 2 (กรุงเทพมหานคร)",
    "ม.5",
    "39",
    "baifernsaeng@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-040",
    "นลพรรณ์ แสงพงษ์พิทยา",
    "บดินทรเดชา 2 (กรุงเทพมหานคร)",
    "ม.5",
    "40",
    "baifernsaeng2@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-041",
    "ธีรภัทร เจริญผิว",
    "โรงเรียนราชสีมาวิทยาลัย (นครราชสีมา)",
    "ม.5",
    "41",
    "xxphoom98xx@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-042",
    "นางสาวณิชากรกร่างสวัสดิ์",
    "โรงเรียนบุญเหลือวิทยานุสรณ์ (นครราชสีมา)",
    "ม.5",
    "42",
    "ksw.mild@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-043",
    "กัลยาณี สุวรรณทา",
    "บุญเหลือวิทยานุสรณ์ (นครราชสีมา)",
    "ม.5",
    "43",
    "Kanlayaneesuwantha7@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-044",
    "นายจักรพล หุนติราช",
    "โรงเรียนเทศบาล๑(บูรพาวิทยากร) (นครราชสีมา)",
    "ม.5",
    "44",
    "Cxllmetulip93@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-045",
    "นภัทร ครุณรัมย์",
    "หนองกี่พิทยาคม (บุรีรัมย์)",
    "ม.5",
    "45",
    "napat210649@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-046",
    "ไอดาภา ยุปานันท์",
    "โรงเรียนธัญบุรี (ปทุมธานี)",
    "ม.5",
    "46",
    "yupananidapha@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-047",
    "น.ส.ชัญญาพัชญ์ โยธา",
    "ระยองวิทยาคม (ระยอง)",
    "ม.5",
    "47",
    "ryw41668@rayongwit.ac.th",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-048",
    "นฤมล สอนสเกตุ",
    "โรงเรียนราชประชาสมาสัยฝ่ายมัธยม รัชดาภิเษก ในพระบรมราชูปถัมภ์ (สมุทรปราการ)",
    "ม.5",
    "48",
    "airf5548@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-049",
    "พัชรินทร์ เรียบร้อย",
    "วีรวัฒน์โยธิน (สุรินทร์)",
    "ม.5",
    "49",
    "Patchxiirin36@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-050",
    "Miss Waratchaya Sopha",
    "โรงเรียนวีรวัฒน์โยธิน (สุรินทร์)",
    "ม.5",
    "50",
    "varthyasopha@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-051",
    "เนตรชนกกระแสงสิงห์",
    "นวมินทราชูทิศ กรุงเทพ (กรุงเทพมหานคร)",
    "ม.6",
    "51",
    "jommotomato@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-052",
    "สุพัตรา บุญรวม",
    "ทวพ (สุพัตรา ทดสอบ) (กาญจนบุรี)",
    "ม.6",
    "52",
    "supattrabnk@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ],
  [
    "STD-053",
    "นางสาวฐิดารัตน์ บุญราศรี",
    "โยธินบำรุง (นครศรีธรรมราช)",
    "ม.6",
    "53",
    "thidaratbunrasri@gmail.com",
    "",
    "Active",
    "2026-10-09 00:00:00"
  ]
];
    if (initialStudents.length > 0) {


  // 0. ชีตบัญชีรายชื่อหลัก (Name_base - ตรวจสอบและสร้างหากยังไม่มี)
  let sheetNameBase = ss.getSheetByName(SHEET_NAME_BASE);
  if (!sheetNameBase) {
    sheetNameBase = ss.insertSheet(SHEET_NAME_BASE);
    const nbHeaders = [
      "รหัสนักเรียน (StudentID)",
      "ชื่อ-นามสกุล (FullName)",
      "โรงเรียน (School)",
      "ระดับชั้น (GradeClass)",
      "เลขที่ (No)",
      "อีเมล (Email)"
    ];
    sheetNameBase.appendRow(nbHeaders);
    formatHeaderRow(sheetNameBase, nbHeaders.length, "#065F46"); // สีเขียวเข้ม Emerald
    // ใส่ข้อมูลตั้งต้น 53 คน
    const nbRows = initialStudents.map(s => [s[0], s[1], s[2], s[3], s[4], s[5]]);
    sheetNameBase.getRange(2, 1, nbRows.length, 6).setValues(nbRows);
  }

  // 1. ชีตฐานข้อมูลนักเรียน (Students_DB)
  let sheetStudents = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheetStudents) {
    sheetStudents = ss.insertSheet(SHEET_STUDENTS);
    const headers = [
      "รหัสนักเรียน (StudentID)",
      "ชื่อ-นามสกุล (FullName)",
      "โรงเรียน (School)",
      "ระดับชั้น/ห้อง (GradeClass)",
      "เลขที่ (No)",
      "อีเมลสำหรับเข้าสู่ระบบ (Email)",
      "รหัสผ่าน (Password)",
      "สถานะบัญชี (Status)",
      "วันที่สร้างบัญชี (CreatedAt)"
    ];
    sheetStudents.appendRow(headers);
    formatHeaderRow(sheetStudents, headers.length, "#065F46"); // สีเขียวเข้ม Emerald

    // นำเข้ารายชื่อนักเรียนตั้งต้น 53 คนจากเอกสารโครงการ (PDF Roster)
    sheetStudents.getRange(2, 1, initialStudents.length, 9).setValues(initialStudents);
    }
  }

  // 1.5 ชีตบันทึกประวัติการเข้าสู่ระบบ (Login_Logs - New Tab Sheet ในไฟล์เดียวกัน)
  let sheetLogins = ss.getSheetByName(SHEET_LOGIN_LOGS);
  if (!sheetLogins) {
    sheetLogins = ss.insertSheet(SHEET_LOGIN_LOGS);
    const loginHeaders = [
      "วันเวลาที่เข้าระบบ (Timestamp)",
      "รหัสนักเรียน (StudentID)",
      "ชื่อ-นามสกุล (FullName)",
      "อีเมล (Email)",
      "โรงเรียน (School)",
      "ระดับชั้น/ห้อง (GradeClass)",
      "ประเภทการเข้า (LoginType)",
      "สถานะการเข้าสู่ระบบ (Status)"
    ];
    sheetLogins.appendRow(loginHeaders);
    formatHeaderRow(sheetLogins, loginHeaders.length, "#0D9488"); // สีเขียวมรกต Teal
  }

  // 2. ชีตบันทึกคะแนน (Scores_Log)
  let sheetScores = ss.getSheetByName(SHEET_SCORES);
  if (!sheetScores) {
    sheetScores = ss.insertSheet(SHEET_SCORES);
    const headers = [
      "วันเวลาที่ส่ง (Timestamp)",
      "รหัสการส่ง (SubmissionID)",
      "อีเมลนักเรียน (Email)",
      "ชื่อ-นามสกุล (FullName)",
      "โรงเรียน (School)",
      "ระดับชั้น/ห้อง (GradeClass)",
      "เลขที่ (No)",
      "บทเรียน (Unit)",
      "คะแนนที่ได้ (Score)",
      "คะแนนเต็ม (TotalScore)",
      "คิดเป็นร้อยละ (%)",
      "สถานะการทำ (Status)",
      "ข้อความถึงคุณครู (StudentNote)",
      "คำแนะนำจากคุณครู (TeacherFeedback)",
      "วันเวลาที่ครูตอบ (FeedbackDate)",
      "คำตอบนักเรียน (AnswersJSON)"
    ];
    sheetScores.appendRow(headers);
    formatHeaderRow(sheetScores, headers.length, "#059669"); // สีเขียวมรกต Forest/Emerald
  }

  // 3. ชีตประวัติเกียรติบัตร (Certificates)
  let sheetCerts = ss.getSheetByName(SHEET_CERTS);
  if (!sheetCerts) {
    sheetCerts = ss.insertSheet(SHEET_CERTS);
    const headers = [
      "รหัสเกียรติบัตร (CertID)",
      "อีเมลนักเรียน (Email)",
      "ชื่อ-นามสกุล (FullName)",
      "โรงเรียน (School)",
      "ระดับชั้น/ห้อง (GradeClass)",
      "บทเรียนที่สำเร็จ (CompletedUnits)",
      "คะแนนเฉลี่ยรวม (%)",
      "วันที่ออกเกียรติบัตร (IssueDate)",
      "สถานะการส่งอีเมล (EmailStatus)",
      "ลิงก์ไฟล์หรือเอกสาร (CertLink)"
    ];
    sheetCerts.appendRow(headers);
    formatHeaderRow(sheetCerts, headers.length, "#047857"); // สีเขียวมรกต Emerald
  }

  // 4. ชีตการตั้งค่าระบบ (Teacher_Config)
  let sheetConfig = ss.getSheetByName(SHEET_CONFIG);
  if (!sheetConfig) {
    sheetConfig = ss.insertSheet(SHEET_CONFIG);
    const headers = ["ชื่อการตั้งค่า (Setting Key)", "ค่าที่บันทึก (Value)", "คำอธิบาย (Description)"];
    sheetConfig.appendRow(headers);
    formatHeaderRow(sheetConfig, headers.length, "#475569"); // สี Slate

    sheetConfig.appendRow(["TEACHER_PIN", "1234", "รหัสผ่านเข้าใช้งาน Dashboard คุณครู"]);
    sheetConfig.appendRow(["ADMIN_EMAIL", Session.getActiveUser().getEmail() || "", "อีเมลคุณครูผู้ดูแลระบบ"]);
    sheetConfig.appendRow(["SYSTEM_VERSION", "v2.0.0", "เวอร์ชันของระบบ"]);
  }

  Logger.log("Setup All Sheets Completed Successfully!");
}

/**
 * ตัวช่วยจัดแต่งแถวหัวตารางให้อ่านง่าย สวยงาม
 */
function formatHeaderRow(sheet, colCount, hexColor) {
  const headerRange = sheet.getRange(1, 1, 1, colCount);
  headerRange.setBackground(hexColor);
  headerRange.setFontColor("#FFFFFF");
  headerRange.setFontWeight("bold");
  headerRange.setHorizontalAlignment("center");
  sheet.setFrozenRows(1);
  for (let c = 1; c <= colCount; c++) {
    sheet.autoResizeColumn(c);
  }
}

/**
 * =========================================================================
 * POST REQUEST ROUTER (รับคำขอส่งข้อมูลจาก WebApp)
 * =========================================================================
 */
function doPost(e) {
  try {
    let payload;
    if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else {
      payload = e.parameter;
    }

    const action = payload.action || "submitScore";
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ---------------------------------------------------------------------
    // 1. ACTION: studentLogin (เข้าสู่ระบบด้วยอีเมลและรหัสผ่าน)
    // ---------------------------------------------------------------------
    if (action === "studentLogin") {
      const email = String(payload.email || "").trim().toLowerCase();
      const password = String(payload.password || "").trim();

      if (!email || !password) {
        return jsonResponse({ status: "error", message: "กรุณาระบุอีเมลและรหัสผ่านให้ครบถ้วน" });
      }

      const sheetStudents = ss.getSheetByName(SHEET_STUDENTS) || setupAllSheets() || ss.getSheetByName(SHEET_STUDENTS);
      const studentData = sheetStudents.getDataRange().getValues();

      let matchedStudent = null;
      for (let i = 1; i < studentData.length; i++) {
        const row = studentData[i];
        const rowEmail = String(row[5] || "").trim().toLowerCase();
        const rowPassword = String(row[6] || "").trim();
        const rowStatus = String(row[7] || "Active").trim();

        if (rowEmail === email) {
          if (rowStatus.toLowerCase() === "suspended" || rowStatus === "ระงับ") {
            return jsonResponse({ status: "error", message: "บัญชีนี้ถูกระงับการใช้งาน กรุณาติดต่อคุณครูผู้สอน" });
          }

          let isFirstTime = false;
          // หากยังไม่มีรหัสผ่านในชีต (นักเรียนลงทะเบียนจาก PDF ครั้งแรก) ให้นำรหัสผ่านที่กรอกมาบันทึกทันที
          if (!rowPassword && password) {
            sheetStudents.getRange(i + 1, 7).setValue(password);
            rowPassword = password;
            isFirstTime = true;
          }

          if (rowPassword === password || (!rowPassword && password === "123456")) {
            matchedStudent = {
              studentId: row[0],
              name: row[1],
              school: row[2],
              gradeClass: row[3],
              studentNo: row[4],
              email: row[5]
            };

            // บันทึกประวัติการเข้าระบบลง New Tab Sheet (Login_Logs) ในไฟล์เดียวกันทันที
            recordLoginLog(ss, matchedStudent, isFirstTime ? "เข้าสู่ระบบครั้งแรก (ตั้งรหัสผ่าน)" : "เข้าสู่ระบบปกติ");
            break;
          }
        }
      }

      if (!matchedStudent) {
        // ตรวจสอบกับชีต Name_base (กรณีมีการเพิ่มหรือแก้ไขรายชื่อใน Name_base บน Google Sheet)
        const nameBaseStudents = getStudentsFromNameBase(ss);
        const nbMatched = nameBaseStudents.find(s => s.email && s.email.toLowerCase() === email);

        if (nbMatched) {
          // เพิ่มนักเรียนจาก Name_base เข้าสู่ Students_DB อัตโนมัติ พร้อมรหัสผ่านที่นักเรียนตั้ง
          const newStudentId = nbMatched.studentId || ("STD-" + Utilities.formatDate(new Date(), "Asia/Bangkok", "yyMM") + "-" + Math.floor(100 + Math.random() * 900));
          const createdAt = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");

          sheetStudents.appendRow([
            newStudentId,
            nbMatched.name,
            nbMatched.school,
            nbMatched.gradeClass,
            nbMatched.studentNo || "-",
            email,
            password,
            "Active",
            createdAt
          ]);

          matchedStudent = {
            studentId: newStudentId,
            name: nbMatched.name,
            school: nbMatched.school,
            gradeClass: nbMatched.gradeClass,
            studentNo: nbMatched.studentNo || "-",
            email: email
          };

          recordLoginLog(ss, matchedStudent, "เข้าสู่ระบบครั้งแรก (ซิงค์จาก Name_base)");
        }
      }

      if (!matchedStudent) {
        return jsonResponse({
          status: "error",
          message: "อีเมลหรือรหัสผ่านไม่ถูกต้อง หรือยังไม่มีรายชื่อในฐานข้อมูล Name_base / Students_DB"
        });
      }

      // ดึงความคืบหน้าของบทเรียนที่นักเรียนคนนี้เคยทำไปแล้ว
      const progress = getStudentUnitProgress(ss, email);

      return jsonResponse({
        status: "success",
        message: "เข้าสู่ระบบสำเร็จ",
        student: matchedStudent,
        progress: progress
      });
    }

    // ---------------------------------------------------------------------
    // 1.5 ACTION: recordLoginLog (บันทึกประวัติการเข้าสู่ระบบลง New Tab Sheet: Login_Logs)
    // ---------------------------------------------------------------------
    if (action === "recordLoginLog" || action === "logAccess") {
      const student = payload.student || {};
      const loginType = payload.loginType || "เข้าสู่ระบบ";
      recordLoginLog(ss, student, loginType);
      return jsonResponse({
        status: "success",
        message: "บันทึกประวัติการเข้าระบบลงแท็บ Login_Logs เรียบร้อยแล้ว",
        timestamp: Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss")
      });
    }

    // ---------------------------------------------------------------------
    // 2. ACTION: studentRegister (ลงทะเบียนนักเรียนใหม่)
    // ---------------------------------------------------------------------
    if (action === "studentRegister") {
      const email = String(payload.email || "").trim().toLowerCase();
      const password = String(payload.password || "").trim();
      const name = String(payload.name || "").trim();
      const school = String(payload.school || "").trim();
      const gradeClass = String(payload.gradeClass || "").trim();
      const studentNo = String(payload.studentNo || "").trim();

      if (!email || !password || !name) {
        return jsonResponse({ status: "error", message: "กรุณากรอกข้อมูลสำคัญ (อีเมล, รหัสผ่าน, ชื่อ-สกุล) ให้ครบถ้วน" });
      }

      const sheetStudents = ss.getSheetByName(SHEET_STUDENTS) || setupAllSheets() || ss.getSheetByName(SHEET_STUDENTS);
      const studentData = sheetStudents.getDataRange().getValues();

      for (let i = 1; i < studentData.length; i++) {
        if (String(studentData[i][5] || "").trim().toLowerCase() === email) {
          return jsonResponse({ status: "error", message: "อีเมลนี้ลงทะเบียนในระบบแล้ว กรุณาเข้าสู่ระบบด้วยรหัสผ่าน" });
        }
      }

      const studentId = "STD-" + Utilities.formatDate(new Date(), "Asia/Bangkok", "yyMM") + "-" + Math.floor(100 + Math.random() * 900);
      const createdAt = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");

      sheetStudents.appendRow([
        studentId,
        name,
        school,
        gradeClass,
        studentNo,
        email,
        password,
        "Active",
        createdAt
      ]);

      return jsonResponse({
        status: "success",
        message: "ลงทะเบียนนักเรียนสำเร็จ",
        student: { studentId, name, school, gradeClass, studentNo, email }
      });
    }

    // ---------------------------------------------------------------------
    // 2.5 ACTION: batchImportStudents (นำเข้ารายชื่อนักเรียนจาก Excel/CSV สู่ระบบ)
    // ---------------------------------------------------------------------
    if (action === "batchImportStudents") {
      const students = Array.isArray(payload.students) ? payload.students : [];
      if (students.length === 0) {
        return jsonResponse({ status: "error", message: "ไม่พบข้อมูลรายชื่อนักเรียนในคำขอ" });
      }

      const sheetStudents = ss.getSheetByName(SHEET_STUDENTS) || setupAllSheets() || ss.getSheetByName(SHEET_STUDENTS);
      const studentData = sheetStudents.getDataRange().getValues();

      // Build email and ID map for fast lookup and updating
      const emailRowMap = {};
      const idRowMap = {};
      for (let i = 1; i < studentData.length; i++) {
        const rowEmail = String(studentData[i][5] || "").trim().toLowerCase();
        const rowId = String(studentData[i][0] || "").trim();
        if (rowEmail) emailRowMap[rowEmail] = i + 1; // 1-based row index
        if (rowId) idRowMap[rowId] = i + 1;
      }

      let insertedCount = 0;
      let updatedCount = 0;
      const nowStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");

      students.forEach(function(s) {
        const studentId = s.studentId || ("STD-" + Math.floor(1000 + Math.random() * 9000));
        const name = s.name || "";
        const school = s.school + (s.province ? (" (" + s.province + ")") : "");
        const gradeClass = s.gradeClass || "ม.3";
        const studentNo = s.studentNo || "";
        const email = String(s.email || "").trim().toLowerCase();
        const password = s.password || "";
        const status = s.status || "Active";
        const createdAt = s.createdAt || nowStr;

        const targetRow = (email && emailRowMap[email]) || (studentId && idRowMap[studentId]);

        if (targetRow) {
          // Update existing row (preserve existing password if new is empty)
          const existingPassword = studentData[targetRow - 1][6];
          const finalPassword = password || existingPassword || "";
          sheetStudents.getRange(targetRow, 1, 1, 9).setValues([[
            studentId, name, school, gradeClass, studentNo, email, finalPassword, status, createdAt
          ]]);
          updatedCount++;
        } else {
          // Append new student row
          sheetStudents.appendRow([
            studentId, name, school, gradeClass, studentNo, email, password, status, createdAt
          ]);
          if (email) emailRowMap[email] = sheetStudents.getLastRow();
          if (studentId) idRowMap[studentId] = sheetStudents.getLastRow();
          insertedCount++;
        }
      });

      return jsonResponse({
        status: "success",
        message: "นำเข้ารายชื่อนักเรียนสำเร็จ (เพิ่มใหม่ " + insertedCount + " คน, อัปเดต " + updatedCount + " คน)",
        insertedCount: insertedCount,
        updatedCount: updatedCount
      });
    }

    // ---------------------------------------------------------------------
    // 3. ACTION: submitScore (บันทึกคะแนน และ ตรวจสอบเงื่อนไขออก E-Certificate)
    // ---------------------------------------------------------------------
    if (action === "submitScore") {
      const sheetScores = ss.getSheetByName(SHEET_SCORES) || setupAllSheets() || ss.getSheetByName(SHEET_SCORES);

      const timestamp = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");
      const submissionId = payload.submissionId || ("SUB-" + Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyyMMdd-HHmmss") + "-" + Math.floor(1000 + Math.random() * 9000));
      const email = String(payload.email || payload.studentEmail || "").trim().toLowerCase();
      const studentName = payload.studentName || "";
      const school = payload.school || "";
      const gradeClass = payload.gradeClass || "";
      const studentNo = payload.studentNo || "";
      const unit = payload.unit || "";
      const score = Number(payload.score) || 0;
      const totalScore = Number(payload.totalScore) || 0;
      const percentage = totalScore > 0 ? ((score / totalScore) * 100).toFixed(1) + "%" : "0%";
      const status = payload.isEarlyExit ? "เลิกทำก่อน" : "ทำครบ 100%";
      const studentNote = payload.studentNote || "";
      const teacherFeedback = payload.teacherFeedback || "";
      const teacherFeedbackDate = payload.teacherFeedbackDate || "";
      const answersJson = typeof payload.answers === "object" ? JSON.stringify(payload.answers) : (payload.answers || "");

      // ป้องกันการส่งซ้ำ
      const data = sheetScores.getDataRange().getValues();
      let existingRow = -1;
      for (let i = 1; i < data.length; i++) {
        if (data[i][1] === submissionId) {
          existingRow = i + 1;
          break;
        }
      }

      const rowData = [
        timestamp,
        submissionId,
        email,
        studentName,
        school,
        gradeClass,
        studentNo,
        unit,
        score,
        totalScore,
        percentage,
        status,
        studentNote,
        teacherFeedback,
        teacherFeedbackDate,
        answersJson
      ];

      if (existingRow > 0) {
        sheetScores.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
      } else {
        sheetScores.appendRow(rowData);
      }

      // ตรวจสอบเงื่อนไขการได้รับ E-Certificate:
      // ต้องทำแบบฝึกหัดครบทั้ง 4 Units (Unit 5, Unit 6, Unit 7, Unit 8) แบบครบสมบูรณ์
      let certificateData = null;
      if (email) {
        certificateData = checkAndIssueCertificate(ss, email, studentName, school, gradeClass);
      }

      return jsonResponse({
        status: "success",
        message: "บันทึกผลคะแนนเรียบร้อยแล้ว",
        submissionId: submissionId,
        timestamp: timestamp,
        certificate: certificateData
      });
    }

    // ---------------------------------------------------------------------
    // 4. ACTION: sendCertificateEmail (ส่งอีเมลแจ้งเตือนเกียรติบัตรให้นักเรียน)
    // ---------------------------------------------------------------------
    if (action === "sendCertificateEmail") {
      const email = String(payload.email || "").trim().toLowerCase();
      const certId = payload.certId || "";
      const studentName = payload.studentName || "";
      const school = payload.school || "";
      const avgScore = payload.avgScore || "";

      if (!email) {
        return jsonResponse({ status: "error", message: "ไม่พบอีเมลผู้รับ" });
      }

      const subject = `🏆 เกียรติบัตรออนไลน์ Get INVOLVED! ม.3 — ${studentName}`;
      const htmlBody = `
        <div style="font-family: 'Sarabun', 'Prompt', Arial, sans-serif; max-width: 650px; margin: 0 auto; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 16px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);">
          <div style="background: linear-gradient(135deg, #065f46, #047857); padding: 32px 24px; text-align: center; color: #ffffff;">
            <h1 style="margin: 0; font-size: 24px; font-weight: 800; letter-spacing: -0.5px;">ยินดีด้วยกับความสำเร็จของคุณ!</h1>
            <p style="margin: 8px 0 0; font-size: 14px; opacity: 0.9;">แบบฝึกหัดภาษาอังกฤษออนไลน์ Get INVOLVED! ม.3 (ทวพ. - TWP)</p>
          </div>
          
          <div style="padding: 32px 24px; color: #1e293b;">
            <p style="font-size: 16px; line-height: 1.6; margin-top: 0;">
              เรียน <strong>${studentName}</strong>,
            </p>
            <p style="font-size: 14px; line-height: 1.7; color: #475569;">
              คุณได้ทำแบบฝึกหัดทบทวนบทเรียนภาษาอังกฤษ ชั้นมัธยมศึกษาปีที่ 1 ครบทุกหน่วยการเรียนรู้ (Unit 5 - 8) สมบูรณ์เรียบร้อยแล้ว ระบบได้ออกเกียรติบัตรอิเล็กทรอนิกส์ (E-Certificate) ประจำตัวคุณเป็นที่เรียบร้อย:
            </p>
            
            <div style="background: #f8fafc; border: 2px dashed #cbd5e1; border-radius: 12px; padding: 20px; margin: 24px 0; text-align: center;">
              <div style="font-size: 12px; font-weight: bold; color: #64748b; text-transform: uppercase;">รหัสเกียรติบัตรอ้างอิง (Serial Number)</div>
              <div style="font-size: 22px; font-weight: 800; color: #047857; font-family: monospace; margin: 6px 0;">${certId}</div>
              <div style="font-size: 13px; color: #334155;">โรงเรียน: <strong>${school}</strong> | คะแนนเฉลี่ยรวม: <strong>${avgScore}</strong></div>
            </div>

            <p style="font-size: 13px; color: #64748b; line-height: 1.6;">
              💡 คุณสามารถเปิดเว็บแอปพลิเคชันเพื่อดูภาพเกียรติบัตรฉบับเต็มและดาวน์โหลดไฟล์ภาพคุณภาพสูงเก็บไว้ในเครื่องได้ตลอดเวลา
            </p>
            
            <div style="margin-top: 32px; padding-top: 20px; border-top: 1px solid #f1f5f9; text-align: center; font-size: 12px; color: #94a3b8;">
              บริษัท สำนักพิมพ์ ไทยวัฒนาพานิช จำกัด (ทวพ.) • โครงการติวเทอมสอง Get INVOLVED! ม.3
            </div>
          </div>
        </div>
      `;

      try {
        MailApp.sendEmail({
          to: email,
          subject: subject,
          htmlBody: htmlBody
        });

        // อัปเดตสถานะในชีต Certificates
        const sheetCerts = ss.getSheetByName(SHEET_CERTS);
        if (sheetCerts) {
          const certsData = sheetCerts.getDataRange().getValues();
          for (let i = 1; i < certsData.length; i++) {
            if (String(certsData[i][0]).trim() === certId || String(certsData[i][1]).trim().toLowerCase() === email) {
              sheetCerts.getRange(i + 1, 9).setValue("ส่งอีเมลแล้ว (Sent: " + Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm") + ")");
              break;
            }
          }
        }

        return jsonResponse({ status: "success", message: `ส่งอีเมลเกียรติบัตรไปยัง ${email} สำเร็จแล้ว` });
      } catch (mailErr) {
        return jsonResponse({ status: "warning", message: `บันทึกข้อมูลเกียรติบัตรแล้ว แต่ไม่สามารถส่งอีเมลได้: ${mailErr.toString()}` });
      }
    }

    // ---------------------------------------------------------------------
    // 5. ACTION: updateTeacherFeedback (ครูส่งข้อเสนอแนะกลับหานักเรียน)
    // ---------------------------------------------------------------------
    if (action === "updateTeacherFeedback") {
      const submissionId = payload.submissionId;
      const feedback = payload.teacherFeedback || "";
      const feedbackDate = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");

      const sheetScores = ss.getSheetByName(SHEET_SCORES);
      if (!sheetScores) return jsonResponse({ status: "error", message: "ไม่พบชีตคะแนน" });

      const data = sheetScores.getDataRange().getValues();
      let foundRow = -1;
      for (let i = 1; i < data.length; i++) {
        if (data[i][1] === submissionId) {
          foundRow = i + 1;
          break;
        }
      }

      if (foundRow > 0) {
        sheetScores.getRange(foundRow, 14).setValue(feedback); // Column N
        sheetScores.getRange(foundRow, 15).setValue(feedbackDate); // Column O
        return jsonResponse({ status: "success", message: "บันทึกข้อเสนอแนะเรียบร้อยแล้ว", feedbackDate: feedbackDate });
      } else {
        return jsonResponse({ status: "error", message: "ไม่พบรหัสการส่งงานนี้ในระบบ" });
      }
    }

    // ---------------------------------------------------------------------
    // 6. ACTION: teacherLogin (ตรวจสอบรหัสผ่านคุณครู)
    // ---------------------------------------------------------------------
    if (action === "teacherLogin") {
      const inputPin = String(payload.pin || "").trim();
      const currentPin = getTeacherPinFromConfig(ss);

      if (inputPin === currentPin) {
        return jsonResponse({ status: "success", message: "เข้าสู่ระบบคุณครูสำเร็จ" });
      } else {
        return jsonResponse({ status: "error", message: "รหัสผ่านคุณครูไม่ถูกต้อง" });
      }
    }

    // ---------------------------------------------------------------------
    // 7. ACTION: updateTeacherPin (เปลี่ยนรหัสผ่านคุณครู)
    // ---------------------------------------------------------------------
    if (action === "updateTeacherPin") {
      const oldPin = String(payload.oldPin || "").trim();
      const newPin = String(payload.newPin || "").trim();
      const currentPin = getTeacherPinFromConfig(ss);

      if (oldPin !== currentPin) {
        return jsonResponse({ status: "error", message: "รหัสผ่านเดิมไม่ถูกต้อง" });
      }
      if (!newPin || newPin.length < 4) {
        return jsonResponse({ status: "error", message: "รหัสผ่านใหม่ต้องมีความยาวอย่างน้อย 4 ตัวอักษร" });
      }

      setTeacherPinInConfig(ss, newPin);
      return jsonResponse({ status: "success", message: "เปลี่ยนรหัสผ่านคุณครูสำเร็จแล้ว" });
    }

    return jsonResponse({ status: "error", message: "ไม่รู้จักคำสั่ง action ที่ระบุ" });

  } catch (error) {
    return jsonResponse({ status: "error", message: error.toString() });
  }
}

/**
 * =========================================================================
 * GET REQUEST ROUTER (เรียกดูข้อมูลเพื่อแสดงผล)
 * =========================================================================
 */
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || "ping";
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. ตรวจสอบการเชื่อมต่อ (Ping)
    if (action === "ping") {
      return jsonResponse({
        status: "ok",
        message: "เชื่อมต่อกับระบบ Google Sheet TWP สำเร็จสมบูรณ์",
        version: "2.0.0",
        time: Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss")
      });
    }

    // ดึงประวัติการเข้าสู่ระบบ (สำหรับแท็บ Login_Logs ใน Dashboard ครู)
    if (action === "getLoginLogs") {
      const sheetLogins = ss.getSheetByName(SHEET_LOGIN_LOGS);
      if (!sheetLogins) return jsonResponse({ status: "success", logs: [] });

      const data = sheetLogins.getDataRange().getValues();
      const logs = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        logs.push({
          timestamp: row[0],
          studentId: row[1],
          studentName: row[2],
          email: row[3],
          school: row[4],
          gradeClass: row[5],
          loginType: row[6],
          status: row[7]
        });
      }
      return jsonResponse({ status: "success", logs: logs });
    }

    // 2. ดึงรายการผลคะแนนทั้งหมด (สำหรับ Dashboard ครู)
    if (action === "getSubmissions") {
      const sheetScores = ss.getSheetByName(SHEET_SCORES);
      if (!sheetScores) return jsonResponse({ status: "success", submissions: [] });

      const data = sheetScores.getDataRange().getValues();
      const rows = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        rows.push({
          timestamp: row[0],
          submissionId: row[1],
          email: row[2],
          studentName: row[3],
          school: row[4],
          gradeClass: row[5],
          studentNo: row[6],
          unit: row[7],
          score: row[8],
          totalScore: row[9],
          percentage: row[10],
          status: row[11],
          studentNote: row[12],
          teacherFeedback: row[13],
          teacherFeedbackDate: row[14],
          answers: row[15]
        });
      }
      return jsonResponse({ status: "success", submissions: rows });
    }

    // 3. ดึงรายชื่อนักเรียนทั้งหมด (ผสานรวมระหว่าง Students_DB และ Name_base)
    if (action === "getStudents") {
      const sheetStudents = ss.getSheetByName(SHEET_STUDENTS);
      const studentMap = {};

      if (sheetStudents) {
        const data = sheetStudents.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const email = String(row[5] || "").trim().toLowerCase();
          if (email) {
            studentMap[email] = {
              studentId: row[0],
              name: row[1],
              school: row[2],
              gradeClass: row[3],
              studentNo: row[4],
              email: row[5],
              hasPassword: !!row[6],
              status: row[7] || "Active",
              createdAt: row[8]
            };
          }
        }
      }

      // ผสานรวมกับข้อมูลล่าสุดจากแท็บ Name_base (หากครูแก้ไขชื่อ/โรงเรียนใน Name_base)
      const nameBaseStudents = getStudentsFromNameBase(ss);
      nameBaseStudents.forEach(nb => {
        const email = (nb.email || "").toLowerCase().trim();
        if (email) {
          if (studentMap[email]) {
            // อัปเดตข้อมูลตาม Name_base ล่าสุด
            studentMap[email].name = nb.name || studentMap[email].name;
            studentMap[email].school = nb.school || studentMap[email].school;
            studentMap[email].gradeClass = nb.gradeClass || studentMap[email].gradeClass;
            if (nb.studentNo) studentMap[email].studentNo = nb.studentNo;
          } else {
            // เพิ่มรายชื่อใหม่ที่เพิ่งเพิ่มใน Name_base
            studentMap[email] = {
              studentId: nb.studentId,
              name: nb.name,
              school: nb.school,
              gradeClass: nb.gradeClass,
              studentNo: nb.studentNo,
              email: nb.email,
              hasPassword: false,
              status: "Active",
              createdAt: "ซิงค์จาก Name_base"
            };
          }
        }
      });

      const students = Object.values(studentMap);
      return jsonResponse({ status: "success", students: students });
    }

    // 4. ดึงรายชื่อเกียรติบัตรที่ออกแล้วทั้งหมด
    if (action === "getCertificates") {
      const sheetCerts = ss.getSheetByName(SHEET_CERTS);
      if (!sheetCerts) return jsonResponse({ status: "success", certificates: [] });

      const data = sheetCerts.getDataRange().getValues();
      const certs = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        certs.push({
          certId: row[0],
          email: row[1],
          studentName: row[2],
          school: row[3],
          gradeClass: row[4],
          completedUnits: row[5],
          averageScore: row[6],
          issueDate: row[7],
          emailStatus: row[8],
          certLink: row[9]
        });
      }
      return jsonResponse({ status: "success", certificates: certs });
    }

    // 5. ตรวจสอบความคืบหน้ารายบุคคลจากอีเมล (Student Progress Check)
    if (action === "checkStudentProgress") {
      const email = String(e.parameter.email || "").trim().toLowerCase();
      if (!email) return jsonResponse({ status: "error", message: "กรุณาระบุอีเมล" });

      const progress = getStudentUnitProgress(ss, email);
      return jsonResponse({ status: "success", email: email, progress: progress });
    }

    return jsonResponse({ status: "error", message: "ไม่รู้จักคำสั่ง action ที่ระบุ" });

  } catch (error) {
    return jsonResponse({ status: "error", message: error.toString() });
  }
}

/**
 * =========================================================================
 * HELPER FUNCTIONS (ฟังก์ชันสนับสนุน)
 * =========================================================================
 */

/**
 * ดึงสถานะความคืบหน้าของแต่ละ Unit ของนักเรียนคนใดคนหนึ่ง
 */
function getStudentUnitProgress(ss, email) {
  const sheetScores = ss.getSheetByName(SHEET_SCORES);
  const unitsCompleted = {
    unit5: false,
    unit6: false,
    unit7: false,
    unit8: false
  };
  const unitScores = {};

  if (sheetScores) {
    const data = sheetScores.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowEmail = String(row[2] || "").trim().toLowerCase();
      const rowUnit = String(row[7] || "");
      const rowPct = parseFloat(row[10]) || 0;
      const rowStatus = String(row[11] || "");

      if (rowEmail === email && rowStatus !== "เลิกทำก่อน") {
        if (rowUnit.includes("Unit 5") || rowUnit.includes("บทที่ 5")) {
          unitsCompleted.unit5 = true;
          unitScores.unit5 = Math.max(unitScores.unit5 || 0, rowPct);
        }
        if (rowUnit.includes("Unit 6") || rowUnit.includes("บทที่ 6")) {
          unitsCompleted.unit6 = true;
          unitScores.unit6 = Math.max(unitScores.unit6 || 0, rowPct);
        }
        if (rowUnit.includes("Unit 7") || rowUnit.includes("บทที่ 7")) {
          unitsCompleted.unit7 = true;
          unitScores.unit7 = Math.max(unitScores.unit7 || 0, rowPct);
        }
        if (rowUnit.includes("Unit 8") || rowUnit.includes("บทที่ 8")) {
          unitsCompleted.unit8 = true;
          unitScores.unit8 = Math.max(unitScores.unit8 || 0, rowPct);
        }
      }
    }
  }

  // ตรวจสอบว่าได้รับเกียรติบัตรหรือยัง
  let certInfo = null;
  const sheetCerts = ss.getSheetByName(SHEET_CERTS);
  if (sheetCerts) {
    const cData = sheetCerts.getDataRange().getValues();
    for (let i = 1; i < cData.length; i++) {
      if (String(cData[i][1]).trim().toLowerCase() === email) {
        certInfo = {
          certId: cData[i][0],
          avgScore: cData[i][6],
          issueDate: cData[i][7],
          emailStatus: cData[i][8]
        };
        break;
      }
    }
  }

  const completedCount = Object.values(unitsCompleted).filter(Boolean).length;

  return {
    units: unitsCompleted,
    scores: unitScores,
    completedCount: completedCount,
    isAllCompleted: completedCount === 4,
    certificate: certInfo
  };
}

/**
 * ตรวจสอบและสร้างรายการในชีต Certificates เมื่อนักเรียนทำครบทุก Unit
 */
function checkAndIssueCertificate(ss, email, studentName, school, gradeClass) {
  const progress = getStudentUnitProgress(ss, email);
  if (!progress.isAllCompleted) return null;

  const sheetCerts = ss.getSheetByName(SHEET_CERTS);
  if (!sheetCerts) return null;

  // ตรวจสอบว่าเคยออกเกียรติบัตรแล้วหรือยัง
  const certData = sheetCerts.getDataRange().getValues();
  for (let i = 1; i < certData.length; i++) {
    if (String(certData[i][1]).trim().toLowerCase() === email) {
      return {
        certId: certData[i][0],
        avgScore: certData[i][6],
        issueDate: certData[i][7],
        isExisting: true
      };
    }
  }

  // คำนวณคะแนนเฉลี่ยรวม
  const scores = Object.values(progress.scores);
  const avg = scores.length > 0 ? (scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(1) + "%" : "100.0%";

  const now = new Date();
  const yearBE = now.getFullYear() + 543;
  const certId = "CERT-TWP-" + yearBE + "-" + Math.floor(1000 + Math.random() * 9000);
  const issueDate = Utilities.formatDate(now, "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");

  const row = [
    certId,
    email,
    studentName,
    school,
    gradeClass,
    "ครบ 4 บทเรียน (Units 5, 6, 7, 8)",
    avg,
    issueDate,
    "พร้อมส่ง (Ready)",
    ""
  ];
  sheetCerts.appendRow(row);

  return {
    certId: certId,
    avgScore: avg,
    issueDate: issueDate,
    isNew: true
  };
}

/**
 * ดึงรหัสผ่านครูจากชีต Teacher_Config (ค่าเริ่มต้น: 1234)
 */
function getTeacherPinFromConfig(ss) {
  const sheetConfig = ss.getSheetByName(SHEET_CONFIG);
  if (!sheetConfig) return "1234";

  const data = sheetConfig.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === "TEACHER_PIN") {
      return String(data[i][1]).trim();
    }
  }
  return "1234";
}

/**
 * บันทึกรหัสผ่านครูใหม่ในชีต Teacher_Config
 */
function setTeacherPinInConfig(ss, newPin) {
  let sheetConfig = ss.getSheetByName(SHEET_CONFIG);
  if (!sheetConfig) {
    setupAllSheets();
    sheetConfig = ss.getSheetByName(SHEET_CONFIG);
  }

  const data = sheetConfig.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === "TEACHER_PIN") {
      sheetConfig.getRange(i + 1, 2).setValue(newPin);
      return;
    }
  }
  sheetConfig.appendRow(["TEACHER_PIN", newPin, "รหัสผ่านเข้าใช้งาน Dashboard คุณครู"]);
}

/**
 * Helper ส่งค่าตอบกลับเป็น JSON Format
 */
/**
 * บันทึกประวัติการเข้าใช้งานระบบของนักเรียนลงในแท็บ Login_Logs (New Tab Sheet ในไฟล์เดียวกัน)
 */
/**
 * ดึงและแมปข้อมูลนักเรียนจากแท็บ Name_base อัตโนมัติ (Dynamic Column Detection)
 * เพื่อใช้เป็นฐานตรวจสอบการเข้าระบบร่วมกับ Students_DB
 */
function getStudentsFromNameBase(ss) {
  const sheetNameBase = ss.getSheetByName(SHEET_NAME_BASE);
  if (!sheetNameBase) return [];

  const data = sheetNameBase.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(h => String(h || "").trim().toLowerCase());
  let colEmail = -1, colName = -1, colSchool = -1, colGrade = -1, colNo = -1, colId = -1;

  headers.forEach((h, idx) => {
    if (colEmail === -1 && (h.includes("email") || h.includes("อีเมล") || h.includes("mail"))) colEmail = idx;
    if (colName === -1 && (h.includes("ชื่อ") || h.includes("name") || h.includes("fullname"))) colName = idx;
    if (colSchool === -1 && (h.includes("โรงเรียน") || h.includes("school") || h.includes("สถานศึกษา"))) colSchool = idx;
    if (colGrade === -1 && (h.includes("ชั้น") || h.includes("grade") || h.includes("ระดับชั้น"))) colGrade = idx;
    if (colNo === -1 && (h.includes("เลขที่") || h === "no" || h.includes("ลำดับ"))) colNo = idx;
    if (colId === -1 && (h.includes("รหัส") || h === "id" || h.includes("studentid"))) colId = idx;
  });

  // Fallback: ตรวจหาคอลัมน์ที่มีเครื่องหมาย @
  if (colEmail === -1) {
    for (let c = 0; c < headers.length; c++) {
      if (data.length > 1 && String(data[1][c]).includes("@")) {
        colEmail = c;
        break;
      }
    }
  }

  const students = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const email = colEmail !== -1 ? String(row[colEmail] || "").trim().toLowerCase() : "";
    const name = colName !== -1 ? String(row[colName] || "").trim() : "";
    const school = colSchool !== -1 ? String(row[colSchool] || "").trim() : "โรงเรียนในโครงการ";
    const gradeClass = colGrade !== -1 ? String(row[colGrade] || "").trim() : "ม.3";
    const studentNo = colNo !== -1 ? String(row[colNo] || "").trim() : String(i);
    const studentId = colId !== -1 ? String(row[colId] || "").trim() : ("STD-" + String(i).padStart(3, "0"));

    if (email || name) {
      students.push({
        studentId: studentId,
        name: name || ("นักเรียนลำดับที่ " + i),
        school: school,
        gradeClass: gradeClass,
        studentNo: studentNo,
        email: email,
        source: "Name_base"
      });
    }
  }
  return students;
}

function recordLoginLog(ss, student, loginType) {
  try {
    let sheetLogins = ss.getSheetByName(SHEET_LOGIN_LOGS);
    if (!sheetLogins) {
      sheetLogins = ss.insertSheet(SHEET_LOGIN_LOGS);
      const loginHeaders = [
        "วันเวลาที่เข้าระบบ (Timestamp)",
        "รหัสนักเรียน (StudentID)",
        "ชื่อ-นามสกุล (FullName)",
        "อีเมล (Email)",
        "โรงเรียน (School)",
        "ระดับชั้น/ห้อง (GradeClass)",
        "ประเภทการเข้า (LoginType)",
        "สถานะการเข้าสู่ระบบ (Status)"
      ];
      sheetLogins.appendRow(loginHeaders);
      formatHeaderRow(sheetLogins, loginHeaders.length, "#0D9488");
    }

    const timestamp = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");
    sheetLogins.appendRow([
      timestamp,
      student.studentId || "-",
      student.name || "-",
      student.email || "-",
      student.school || "-",
      student.gradeClass || "-",
      loginType || "เข้าสู่ระบบ",
      "สำเร็จ (Online)"
    ]);
  } catch (err) {
    Logger.log("Error recording login log: " + err.toString());
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
