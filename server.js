import express from "express";
import multer from "multer";
import XLSX from "xlsx";
import PDFDocument from "pdfkit";
import archiver from "archiver";
import QRCode from "qrcode";
import bcrypt from "bcrypt";
import { v4 as uuidv4 } from "uuid";
import fs from "fs";

const app = express();
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use("/generated", express.static("generated"));

if (!fs.existsSync("generated")) fs.mkdirSync("generated");
if (!fs.existsSync("uploads")) fs.mkdirSync("uploads");

const upload = multer({ dest: "uploads/" });

// ================== ADMIN LOGIN ==================
const ADMIN_USERNAME = "admin";
const HASHED_PASSWORD = bcrypt.hashSync("admin123", 10);

// ================== SETTINGS ==================
let SETTINGS = {
  schoolName: "RUBONGI ARMY SECONDARY SCHOOL",
  address: "P.O.BOX 698 TORORO, TEL:0454445148/0782651148",
  vision: "To produce a morally upright and self reliant future generation.",
  mission: "To provide affordable quality education to our community.",
  footer: "Victory Is Our Challenge",
  headTeacher: "ZAINA .K. NALUKENGE",
  headTeacherRank: "Maj.",
  headTeacherTitle: "HEAD TEACHER"
};

let LOGO1 = null;
let LOGO2 = null;
let DATABASE = {};
let serialCounter = 1;

// ================== UACE SUBJECT MAPPING & SUBSIDIARY CODES ==================
const SUBJECT_NAMES = {
  ENG: "ENGLISH",
  HIS: "HISTORY",
  GEO: "GEOGRAPHY",
  MAT: "MATHEMATICS",
  PHY: "PHYSICS",
  CHE: "CHEMISTRY",
  BIO: "BIOLOGY",
  IPS: "IPS",
  CRE: "CRE",
  COM: "COMMERCE",
  IRE: "IRE",
  AGR: "AGRICULTURE",
  DHP: "DHOPADHOLA",
  LIT: "LITERATURE IN ENGLISH",
  ENT: "ENTREPRENEURSHIP",
  KIS: "KISWAHILI",
  LAN: "LANGO",
  PE:  "PHYSICAL EDUCATION",
  PA:  "PERFORMING ARTS",
  FRE: "FRENCH",
  ECO: "ECONOMICS",
  // Subsidiaries
  GEP: "GENERAL PAPER",
  CST: "ICT (SUBSIDIARY)",
  SMA: "SUB-MATHEMATICS (SUBSIDIARY)"
};

const SUBSIDIARY_CODES = new Set(['GEP', 'CST', 'SMA']);

// Grade points for principle subjects
const GRADE_POINTS = { 'A':6, 'B':5, 'C':4, 'D':3, 'E':2, 'O':1, 'F':0 };

// ================== UACE PARSING FUNCTIONS ==================
function numericToLetterGrade(num) {
  const map = { 1:'A', 2:'B', 3:'C', 4:'D', 5:'E', 6:'O', 7:'F', 8:'F', 9:'F' };
  return map[num] || 'F';
}

function parseOverallGrade(token, isSubsidiary) {
  if (!token || token.trim() === '') return isSubsidiary ? 9 : 'F';
  const trimmed = token.trim();
  if (isSubsidiary) {
    const num = parseInt(trimmed, 10);
    return (!isNaN(num) && num >= 1 && num <= 9) ? num : 9;
  } else {
    const first = trimmed.charAt(0).toUpperCase();
    if (['A','B','C','D','E','F','O'].includes(first)) return first;
    const num = parseInt(trimmed, 10);
    if (!isNaN(num) && num >= 1 && num <= 9) return numericToLetterGrade(num);
    return 'F';
  }
}

function parseSubjectEntry(entry) {
  // Format: "CODE-GRADE [1-6,2-7,3-7]"
  const regex = /^([A-Za-z]+)-([A-Z0-9]+)\s*\[(.*?)\]$/;
  const match = entry.match(regex);
  if (!match) return null;
  const code = match[1];
  const overallToken = match[2];
  const paperStr = match[3];
  const isSub = SUBSIDIARY_CODES.has(code);
  const overallGrade = parseOverallGrade(overallToken, isSub);

  const paperGrades = [];
  if (paperStr.trim() !== '') {
    const parts = paperStr.split(',');
    parts.forEach(p => {
      const pair = p.split('-');
      if (pair.length === 2) {
        const paperNum = parseInt(pair[0].trim(), 10);
        const paperGrade = parseInt(pair[1].trim(), 10);
        if (!isNaN(paperNum) && !isNaN(paperGrade) && paperGrade >=1 && paperGrade <=9) {
          paperGrades.push({ paper: paperNum, grade: paperGrade });
        }
      }
    });
  }
  return { code, overallGrade, isSubsidiary: isSub, paperGrades };
}

// ================== LOGIN PAGE ==================
app.get("/", (req, res) => {
  res.send(`
  <!DOCTYPE html>
  <html>
  <head>
    <title>UACE Testimonial System</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
      * { box-sizing: border-box; margin: 0; padding: 0; }
      body {
        font-family: 'Segoe UI', Roboto, sans-serif;
        background: linear-gradient(135deg, #1b4f6e, #123a4f);
        min-height: 100vh;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 20px;
      }
      .card {
        background: white;
        border-radius: 20px;
        padding: 40px;
        width: 100%;
        max-width: 420px;
        box-shadow: 0 20px 60px rgba(0,0,0,0.3);
      }
      h2 { color: #1b4f6e; margin-bottom: 30px; text-align: center; }
      input {
        width: 100%;
        padding: 15px;
        margin: 10px 0;
        border: 2px solid #e0e0e0;
        border-radius: 10px;
        font-size: 16px;
      }
      input:focus { border-color: #1b4f6e; outline: none; }
      button {
        width: 100%;
        padding: 15px;
        background: #1b4f6e;
        color: white;
        border: none;
        border-radius: 10px;
        font-size: 16px;
        font-weight: 600;
        cursor: pointer;
        margin-top: 20px;
      }
      button:hover { background: #123a4f; }
    </style>
  </head>
  <body>
    <div class="card">
      <h2>🔐 UACE Admin Login</h2>
      <form method="POST" action="/dashboard">
        <input name="username" placeholder="Username" required autofocus>
        <input name="password" type="password" placeholder="Password" required>
        <button>Login</button>
      </form>
    </div>
  </body>
  </html>
  `);
});

app.post("/dashboard", async (req, res) => {
  if (req.body.username !== ADMIN_USERNAME) return res.send("Invalid login");
  const valid = await bcrypt.compare(req.body.password, HASHED_PASSWORD);
  if (!valid) return res.send("Invalid login");
  res.send(DASHBOARD_HTML());
});

function DASHBOARD_HTML() {
  return `
  <!DOCTYPE html>
  <html>
  <head>
    <title>UACE Dashboard</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
      * { box-sizing: border-box; margin: 0; padding: 0; }
      body {
        font-family: 'Segoe UI', Roboto, sans-serif;
        background: #f0f5f9;
        padding: 30px 20px;
      }
      .container { max-width: 900px; margin: 0 auto; }
      h1 {
        color: #1b4f6e;
        margin-bottom: 30px;
        border-left: 8px solid #1b4f6e;
        padding-left: 20px;
      }
      .box {
        background: white;
        border-radius: 16px;
        padding: 30px;
        margin-bottom: 30px;
        box-shadow: 0 5px 20px rgba(0,0,0,0.05);
      }
      h3 { color: #1b4f6e; margin-bottom: 20px; }
      input, textarea, button, input[type="file"] {
        width: 100%;
        padding: 12px 16px;
        margin: 8px 0;
        border: 2px solid #d9e2e9;
        border-radius: 10px;
        font-size: 1rem;
        font-family: inherit;
      }
      input:focus, textarea:focus { border-color: #1b4f6e; outline: none; }
      button {
        background: #1b4f6e;
        color: white;
        font-weight: 600;
        border: none;
        cursor: pointer;
        margin-top: 15px;
      }
      button:hover { background: #123a4f; }
      .logout { text-align: right; margin-bottom: 20px; }
      .logout a { color: #c44545; text-decoration: none; }
      .footer {
        text-align: center;
        margin-top: 40px;
        padding: 20px;
        background: white;
        border-radius: 16px;
        color: #1b4f6e;
      }
    </style>
  </head>
  <body>
    <div class="container">
      <div class="logout"><a href="/">← Logout</a></div>
      <h1>📋 UACE Testimonial Generator</h1>

      <div class="box">
        <h3>🖼️ Upload Logos</h3>
        <form action="/upload-assets" method="POST" enctype="multipart/form-data">
          <input type="file" name="logo1" accept="image/*" required>
          <input type="file" name="logo2" accept="image/*" required>
          <button>Upload Logos</button>
        </form>
      </div>

      <div class="box">
        <h3>⚙️ School Settings</h3>
        <form method="POST" action="/settings">
          <input name="schoolName" placeholder="School Name" value="${SETTINGS.schoolName}">
          <input name="address" placeholder="Address & Phone" value="${SETTINGS.address}">
          <textarea name="vision" placeholder="Vision">${SETTINGS.vision}</textarea>
          <textarea name="mission" placeholder="Mission">${SETTINGS.mission}</textarea>
          <input name="footer" placeholder="Footer Motto" value="${SETTINGS.footer}">
          <input name="headTeacher" placeholder="Head Teacher Name" value="${SETTINGS.headTeacher}">
          <input name="headTeacherRank" placeholder="Head Teacher Rank" value="${SETTINGS.headTeacherRank}">
          <input name="headTeacherTitle" placeholder="Head Teacher Title" value="${SETTINGS.headTeacherTitle}">
          <button>Save Settings</button>
        </form>
      </div>

      <div class="box">
        <h3>📊 Generate UACE Testimonials</h3>
        <p><strong>Excel columns (as in sample):</strong> IndexNo, Sex, Candidate_Name, Res. Code, DATE OF BIRTH, Subjects</p>
        <p><strong>Subjects format:</strong> e.g., <code>GEP-5 [1-5]  ENT-O [1-9,2-8,3-5]  CRE-O [1-4,2-5,4-9]</code></p>
        <form action="/generate" method="POST" enctype="multipart/form-data">
          <input type="file" name="excel" accept=".xlsx, .xls, .csv" required>
          <button>Generate ZIP with PDFs</button>
        </form>
      </div>

      <div class="footer">
        <p><strong>Mawerere Francis</strong> · 0788223215 · mawererefrancis@gmail.com</p>
      </div>
    </div>
  </body>
  </html>
  `;
}

// ================== ASSETS ==================
app.post("/upload-assets", upload.fields([
  { name: "logo1" }, { name: "logo2" }
]), (req, res) => {
  LOGO1 = req.files.logo1[0].path;
  LOGO2 = req.files.logo2[0].path;
  res.send("✅ Logos uploaded. <a href='/dashboard'>Back</a>");
});

// ================== SETTINGS ==================
app.post("/settings", (req, res) => {
  SETTINGS = { ...SETTINGS, ...req.body };
  res.send("✅ Settings updated. <a href='/dashboard'>Back</a>");
});

// ================== GENERATE UACE TESTIMONIALS ==================
app.post("/generate", upload.single("excel"), async (req, res) => {
  try {
    const workbook = XLSX.readFile(req.file.path);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const students = XLSX.utils.sheet_to_json(sheet, { defval: '' });

    const zipPath = "generated/uace_testimonials.zip";
    const output = fs.createWriteStream(zipPath);
    const archive = archiver("zip");
    archive.pipe(output);

    serialCounter = 1;

    for (const s of students) {
      // Read columns based on the sample headers
      const name = s["Candidate_Name"] || "";
      const indexNo = s["IndexNo"] || "";
      const sex = s["Sex"] || "";
      const resCode = s["Res. Code"] || "";               // Directly from Excel
      const dob = s["DATE OF BIRTH"] || "";               // Correct column header
      const subjectsStr = s["Subjects"] || "";

      // ---------- IMPROVED TOKENIZATION ----------
      // Split on whitespace, then combine any token starting with '[' with the previous token.
      const parts = subjectsStr.toString().split(/\s+/);
      const combinedTokens = [];
      for (let i = 0; i < parts.length; i++) {
        if (parts[i].startsWith('[')) {
          if (combinedTokens.length > 0) {
            combinedTokens[combinedTokens.length - 1] += ' ' + parts[i];
          }
        } else {
          combinedTokens.push(parts[i]);
        }
      }

      const subjectDetails = [];
      combinedTokens.forEach(token => {
        if (!token.includes('-') || !token.includes('[')) return;
        const parsed = parseSubjectEntry(token);
        if (parsed) {
          parsed.fullName = SUBJECT_NAMES[parsed.code] || parsed.code;
          subjectDetails.push(parsed);
        }
      });

      // Determine maximum number of papers for this candidate
      const maxPapers = Math.max(...subjectDetails.map(s => s.paperGrades.length), 0);

      const gender = sex === "M" ? "MALE" : sex === "F" ? "FEMALE" : sex;
      const genderCode = sex === "M" ? "M" : sex === "F" ? "F" : "X";
      const serialNumber = `UNEB/UACE/${genderCode}/${String(serialCounter).padStart(3, '0')}/2025`;
      serialCounter++;

      // ----- Compute result statistics -----
      let principalPasses = 0;      // A-E in principle subjects
      let subsidiaryPasses = 0;     // 1-6 in subsidiary + O in principle
      let totalPoints = 0;

      subjectDetails.forEach(subj => {
        if (subj.isSubsidiary) {
          // subsidiary subject
          const grade = subj.overallGrade; // numeric 1-9
          if (grade >= 1 && grade <= 6) {
            subsidiaryPasses += 1;
            totalPoints += 1; // subsidiary pass gives 1 point
          }
          // else 0 points
        } else {
          // principle subject
          const grade = subj.overallGrade; // letter
          if (grade === 'O') {
            subsidiaryPasses += 1; // counts as subsidiary pass
            totalPoints += 1; // O gives 1 point
          } else if (['A','B','C','D','E'].includes(grade)) {
            principalPasses += 1;
            totalPoints += GRADE_POINTS[grade] || 0;
          } else {
            // F gives 0 points
            totalPoints += 0;
          }
        }
      });

      const id = uuidv4();
      DATABASE[id] = {
        name, indexNo, sex: gender, dob, year: "2025", serialNumber,
        resCode,                          // store the Res. Code from Excel
        subjectDetails,
        principalPasses,
        subsidiaryPasses,
        totalPoints
      };

      // QR data now includes school name, Res. Code, and computed stats
      const qrData = JSON.stringify({
        school: SETTINGS.schoolName,
        name, indexNo, sex: gender, dob, year: "2025", serialNumber,
        resCode,
        principalPasses,
        subsidiaryPasses,
        totalPoints,
        subjects: subjectDetails.map(s => ({
          code: s.code,
          name: s.fullName,
          grade: s.overallGrade,
          papers: s.paperGrades
        }))
      });
      const qrImage = await QRCode.toDataURL(qrData);

      const safeName = name.replace(/[^a-z0-9]/gi, "_").substring(0, 50);
      const filePath = `generated/${safeName}.pdf`;

      const doc = new PDFDocument({ size: "A4", margin: 40 });
      const writeStream = fs.createWriteStream(filePath);
      doc.pipe(writeStream);

      // ---------- THICK, PROFESSIONAL BORDER ----------
      const borderMargin = 12;
      const borderWidth = doc.page.width - 2 * borderMargin;
      const borderHeight = doc.page.height - 2 * borderMargin;
      const cornerRadius = 20;
      // Outer border (thick dark blue)
      doc.roundedRect(borderMargin, borderMargin, borderWidth, borderHeight, cornerRadius)
         .lineWidth(4).strokeColor("#1b4f6e").stroke();
      // Inner border (thinner gold accent)
      doc.roundedRect(borderMargin + 4, borderMargin + 4, borderWidth - 8, borderHeight - 8, cornerRadius)
         .lineWidth(2).strokeColor("#c9a959").stroke();

      // ---------- SERIAL NUMBER ----------
      doc.fontSize(8).fillColor("#1b4f6e").text(serialNumber, 45, 45);

      // ---------- LOGOS ----------
      const titleY = 170;
      const logoHeight = 70;
      if (LOGO1 && fs.existsSync(LOGO1)) {
        doc.image(LOGO1, 45, titleY - logoHeight, { height: logoHeight });
      }
      if (LOGO2 && fs.existsSync(LOGO2)) {
        doc.image(LOGO2, doc.page.width - 115, titleY - logoHeight, { height: logoHeight });
      }

      // School header
      doc.fontSize(16).fillColor("#1b4f6e").text(SETTINGS.schoolName, 0, 90, { align: "center" });
      doc.fontSize(9).fillColor("#2d3748").text(SETTINGS.address, { align: "center" });
      doc.fontSize(9).fillColor("#4a5568").text(`VISION: ${SETTINGS.vision}`, { align: "center" });
      doc.fontSize(9).text(`MISSION: ${SETTINGS.mission}`, { align: "center" });

      // Title
      doc.fontSize(14).fillColor("#1b4f6e").text("UACE TESTIMONIAL 2025", 0, titleY, { align: "center", underline: true });

      // ---------- CANDIDATE DETAILS BOX ----------
      const boxLeft = 45;
      const boxWidth = 520;
      const boxPadding = 10;
      const nameLength = name.length;
      const boxHeight = nameLength > 35 ? 120 : 100;
      const boxTop = 210;

      doc.roundedRect(boxLeft, boxTop, boxWidth, boxHeight, 5).lineWidth(1.5).strokeColor("#1b4f6e").stroke();

      doc.fontSize(11).fillColor("black");
      if (nameLength > 35) {
        const nameParts = name.split(' ');
        const midPoint = Math.ceil(nameParts.length / 2);
        const line1 = nameParts.slice(0, midPoint).join(' ');
        const line2 = nameParts.slice(midPoint).join(' ');
        doc.text(`CANDIDATE'S NAME: ${line1}`, boxLeft + boxPadding, boxTop + boxPadding);
        doc.text(`${line2}`, boxLeft + boxPadding + 120, boxTop + boxPadding + 18);
        doc.text(`INDEX NO: ${indexNo}`, boxLeft + 300, boxTop + boxPadding);
        doc.text(`SEX: ${gender}`, boxLeft + boxPadding, boxTop + boxPadding + 40);
        doc.text(`DoB: ${dob}`, boxLeft + 200, boxTop + boxPadding + 40);
        doc.text("LIN............................................", boxLeft + boxPadding, boxTop + boxPadding + 60);
      } else {
        doc.text(`CANDIDATE'S NAME: ${name}`, boxLeft + boxPadding, boxTop + boxPadding);
        doc.text(`INDEX NO: ${indexNo}`, boxLeft + 300, boxTop + boxPadding);
        doc.text(`SEX: ${gender}`, boxLeft + boxPadding, boxTop + boxPadding + 20);
        doc.text(`DoB: ${dob}`, boxLeft + 200, boxTop + boxPadding + 20);
        doc.text("LIN............................................", boxLeft + boxPadding, boxTop + boxPadding + 40);
      }

      // ---------- SUBJECT TABLE (Subjects as rows, Papers as columns) ----------
      const tableTop = boxTop + boxHeight + 25;
      const colSubject = 50;            // start of Subject column
      const colFirstPaper = 200;         // start of first paper column
      const paperColWidth = 35;          // width per paper column
      const colOverall = colFirstPaper + maxPapers * paperColWidth + 10; // Overall grade column
      const rowHeight = 28;
      let y = tableTop;

      // Table header background
      doc.rect(colSubject - 2, y - 2, (colOverall + 60) - colSubject + 4, rowHeight)
         .fillColor("#1b4f6e").fill();

      doc.fillColor("white").font("Helvetica-Bold").fontSize(10);
      doc.text("SUBJECT", colSubject + 5, y + 8);
      for (let i = 1; i <= maxPapers; i++) {
        doc.text(`P${i}`, colFirstPaper + (i-1)*paperColWidth + 8, y + 8);
      }
      doc.text("OVERALL", colOverall + 5, y + 8);
      y += rowHeight;

      // Reset fill color for rows
      doc.fillColor("black");

      // Draw vertical lines (thick)
      doc.lineWidth(1.5).strokeColor("#1b4f6e");
      doc.moveTo(colSubject, tableTop).lineTo(colSubject, y + subjectDetails.length * rowHeight + 5).stroke();
      for (let i = 0; i <= maxPapers; i++) {
        const x = colFirstPaper + i * paperColWidth;
        doc.moveTo(x, tableTop).lineTo(x, y + subjectDetails.length * rowHeight + 5).stroke();
      }
      doc.moveTo(colOverall, tableTop).lineTo(colOverall, y + subjectDetails.length * rowHeight + 5).stroke();
      doc.moveTo(colOverall + 60, tableTop).lineTo(colOverall + 60, y + subjectDetails.length * rowHeight + 5).stroke();

      // Horizontal header line
      doc.moveTo(colSubject, tableTop + rowHeight).lineTo(colOverall + 60, tableTop + rowHeight).stroke();

      // Data rows
      subjectDetails.forEach((subj, idx) => {
        // Subject name (may wrap)
        doc.font("Helvetica").fontSize(10);
        doc.text(subj.fullName, colSubject + 5, y + 5, { width: colFirstPaper - colSubject - 10 });

        // Paper grades
        for (let i = 1; i <= maxPapers; i++) {
          const paper = subj.paperGrades.find(p => p.paper === i);
          const gradeText = paper ? paper.grade.toString() : "";
          doc.text(gradeText, colFirstPaper + (i-1)*paperColWidth + 12, y + 8, { width: paperColWidth - 5, align: "center" });
        }

        // Overall grade (big & bold)
        doc.font("Helvetica-Bold").fontSize(16).fillColor("#1b4f6e");
        doc.text(subj.overallGrade.toString(), colOverall + 15, y, { width: 50 });

        y += rowHeight;
        doc.fillColor("black");
      });

      // Horizontal lines between rows
      doc.lineWidth(0.5).strokeColor("#cccccc");
      for (let i = 0; i <= subjectDetails.length; i++) {
        const lineY = tableTop + rowHeight + i * rowHeight;
        doc.moveTo(colSubject, lineY).lineTo(colOverall + 60, lineY).stroke();
      }

      // ---------- RESULT STATISTICS BOX (Res. Code, Principal Passes, Subsidiary Passes, Total Points) ----------
      const statsY = y + 20;
      const statsBoxX = 45;
      const statsBoxWidth = 520;
      const statsBoxHeight = 60;
      doc.roundedRect(statsBoxX, statsY, statsBoxWidth, statsBoxHeight, 5)
         .lineWidth(1.5).strokeColor("#1b4f6e").stroke();

      doc.fontSize(10).font("Helvetica").fillColor("black");
      doc.text(`Res. Code: ${resCode}`, statsBoxX + 10, statsY + 10);
      doc.text(`Principal Passes: ${principalPasses}`, statsBoxX + 200, statsY + 10);
      doc.text(`Subsidiary Passes: ${subsidiaryPasses}`, statsBoxX + 350, statsY + 10);
      doc.text(`Total Points: ${totalPoints}`, statsBoxX + 10, statsY + 35);

      // ---------- MOTTO ----------
      const mottoY = statsY + statsBoxHeight + 15;
      doc.fontSize(11).font("Helvetica-Oblique").fillColor("#1b4f6e")
         .text(SETTINGS.footer, 50, mottoY, { align: "center" });

      // ---------- SIGNATURE BLOCK ----------
      const sigY = mottoY + 40;
      const sigX = 350;
      doc.fontSize(11).font("Helvetica").fillColor("black");
      doc.text("....................................", sigX, sigY - 10);
      doc.text(SETTINGS.headTeacher, sigX, sigY);
      doc.text(SETTINGS.headTeacherRank, sigX, sigY + 18);
      doc.text(SETTINGS.headTeacherTitle, sigX, sigY + 36);

      // ---------- QR CODE ----------
      const qrY = sigY + 70;
      doc.image(qrImage, 45, qrY, { width: 80 });

      doc.end();

      await new Promise(resolve => writeStream.on("finish", resolve));
      archive.file(filePath, { name: `${safeName}.pdf` });
    }

    await archive.finalize();
    output.on("close", () => res.download(zipPath, "uace_testimonials.zip"));
  } catch (error) {
    console.error(error);
    res.status(500).send("Error: " + error.message);
  }
});

// ================== VERIFICATION ==================
app.get("/verify/:id", (req, res) => {
  const s = DATABASE[req.params.id];
  if (!s) return res.send("<h2>Invalid Certificate</h2>");
  let subjectsHtml = '';
  if (s.subjectDetails) {
    subjectsHtml = '<h3>Subjects</h3><table border="1" cellpadding="5" style="border-collapse: collapse; width:100%;"><tr><th>Subject</th><th>Overall Grade</th><th>Paper Grades</th></tr>';
    s.subjectDetails.forEach(subj => {
      const papers = subj.paperGrades.map(p => `P${p.paper}:${p.grade}`).join(', ');
      subjectsHtml += `<tr><td>${subj.fullName || subj.code}</td><td><strong>${subj.overallGrade}</strong></td><td>${papers}</td></tr>`;
    });
    subjectsHtml += '</table>';
  }
  res.send(`
  <!DOCTYPE html>
  <html>
  <head>
    <title>UACE Certificate Verification</title>
    <style>
      body{font-family:Arial;background:#f4f6f9;padding:40px;}
      .card{background:white;padding:30px;border-radius:10px;max-width:700px;margin:auto;box-shadow:0 5px 20px rgba(0,0,0,0.1);}
      h2{color:#1b4f6e;}
      .valid{color:green;font-weight:bold;font-size:24px;}
      .info-grid{display:grid;grid-template-columns:1fr 2fr;gap:10px;margin:20px 0;}
      .label{font-weight:bold;color:#555;}
      table{width:100%; border-collapse: collapse; margin-top:20px;}
      th{background:#1b4f6e; color:white; padding:8px;}
      td{padding:8px; border:1px solid #ccc;}
      .stats { background: #eef7ff; padding: 10px; border-radius: 8px; margin-top: 20px; }
    </style>
  </head>
  <body>
    <div class="card">
      <h2>✅ UACE Certificate Verification</h2>
      <div class="info-grid">
        <div class="label">Serial Number:</div><div>${s.serialNumber || 'N/A'}</div>
        <div class="label">Candidate's Name:</div><div>${s.name}</div>
        <div class="label">Index Number:</div><div>${s.indexNo}</div>
        <div class="label">Sex:</div><div>${s.sex || 'N/A'}</div>
        <div class="label">Date of Birth:</div><div>${s.dob || 'N/A'}</div>
        <div class="label">Year:</div><div>${s.year || '2025'}</div>
        <div class="label">Res. Code:</div><div>${s.resCode || 'N/A'}</div>
        <div class="label">Principal Passes:</div><div>${s.principalPasses || 0}</div>
        <div class="label">Subsidiary Passes:</div><div>${s.subsidiaryPasses || 0}</div>
        <div class="label">Total Points:</div><div>${s.totalPoints || 0}</div>
      </div>
      ${subjectsHtml}
      <h3 class="valid">STATUS: VALID</h3>
    </div>
  </body>
  </html>
  `);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`✅ UACE Server running on http://localhost:${PORT}`));
