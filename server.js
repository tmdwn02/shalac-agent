require('dotenv').config();
const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const { google } = require('googleapis');
const OpenAI = require('openai');
const nodemailer = require('nodemailer');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const multer = require('multer');
const https = require('https');
const http  = require('http');
const pdfParse = require('pdf-parse');

// multer: 메모리에 임시 저장 (디스크 저장 불필요)
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024 }, // 20MB 제한
});

const app = express();
app.use(cors());
app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, 'public')));

// ─── OpenAI ──────────────────────────────────────────────────────────────────
const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

// ─── Google Sheets/Drive Auth ─────────────────────────────────────────────────
const SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive',
  'https://www.googleapis.com/auth/calendar',
];

async function getAuthClient() {
  // 환경변수 GOOGLE_CREDENTIALS_JSON 이 있으면 인라인 JSON으로 인증 (배포 환경)
  // 없으면 GOOGLE_APPLICATION_CREDENTIALS 파일 경로로 인증 (로컬 환경)
  const auth = process.env.GOOGLE_CREDENTIALS_JSON
    ? new google.auth.GoogleAuth({
        credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON),
        scopes: SCOPES,
      })
    : new google.auth.GoogleAuth({
        keyFile: process.env.GOOGLE_APPLICATION_CREDENTIALS,
        scopes: SCOPES,
      });
  return auth.getClient();
}

async function getSheetsClient() {
  const auth = await getAuthClient();
  return google.sheets({ version: 'v4', auth });
}

// ─── 스프레드시트 ID 설정 ──────────────────────────────────────────────────────
const SPREADSHEET_IDS = {
  members: process.env.SHEET_MEMBERS_ID || '',
  guests: process.env.SHEET_GUESTS_ID || '',
  training: process.env.SHEET_TRAINING_ID || '',
  budget: process.env.SHEET_BUDGET_ID || '',
  equipment: process.env.SHEET_EQUIPMENT_ID || '',
  queries: process.env.SHEET_QUERIES_ID || '',
};

// ─── 동아리 운영 메뉴얼 파일 목록 ─────────────────────────────────────────────
const MANUAL_DOCS = [
  { id: '1cEGmm0VtMwMjrHvsz8sop7EV4Fz-npiETD0V0lzUfFI', name: 'SNS관리 메뉴얼' },
  { id: '1FiZzUiI__HpfcERrAdvflDTFB7IqPh_9UbCLqJOv4l8', name: '정기 연습 & 대회 메뉴얼' },
  { id: '1-aYRqC6nVguWRZ1GFM2SpW97Sz3iKNOZ-0AA9HUfBms', name: '동소제 매뉴얼' },
  { id: '1j-NFcTr0WU-tg9vcj69SCuMu6LMZb42guvUnT3DAYII', name: '부서 & 행사 메뉴얼' },
  { id: '1uafo9XIxslSt-1yQoEyZ0MKQ34yLLLj_LICfqYtm2gk', name: '신입부원 모집 메뉴얼' },
  { id: '1Fx4CD9Gvria4cw7arFFhxwlIkQsSd2eW1w2guIsNoYg',  name: '장비 메뉴얼' },
  { id: '1Dgieu8UL5w-U1qE4A_2hcQRdEj1Ys6yGQRJ7nROppX8', name: '동아리연합회 관련 메뉴얼' },
];

// 메뉴얼 캐시 (서버 시작 시 1회 로드)
let manualCache = null;

async function loadManuals() {
  try {
    const auth = await getAuthClient();
    const drive = google.drive({ version: 'v3', auth });
    const docs = {};
    for (const doc of MANUAL_DOCS) {
      const res = await drive.files.export(
        { fileId: doc.id, mimeType: 'text/plain' },
        { responseType: 'text' }
      );
      // BOM 제거 + 과도한 공백 정리
      docs[doc.name] = res.data.replace(/^﻿/, '').replace(/\n{3,}/g, '\n\n').trim();
      console.log(`✅ 메뉴얼 로드: ${doc.name} (${docs[doc.name].length}자)`);
    }
    // ─── 로컬 manuals/ 폴더의 PDF/TXT 파일도 로드 ───────────────────────────
    const manualsDir = path.join(__dirname, 'manuals');
    if (fs.existsSync(manualsDir)) {
      const files = fs.readdirSync(manualsDir);
      for (const file of files) {
        const filePath = path.join(manualsDir, file);
        const name = path.basename(file, path.extname(file));
        try {
          if (file.endsWith('.pdf')) {
            const buffer = fs.readFileSync(filePath);
            const data = await pdfParse(buffer);
            docs[name] = data.text.replace(/\n{3,}/g, '\n\n').trim();
            console.log(`✅ PDF 메뉴얼 로드: ${name} (${docs[name].length}자)`);
          } else if (file.endsWith('.txt')) {
            docs[name] = fs.readFileSync(filePath, 'utf-8').replace(/\n{3,}/g, '\n\n').trim();
            console.log(`✅ TXT 메뉴얼 로드: ${name} (${docs[name].length}자)`);
          }
        } catch (fileErr) {
          console.error(`❌ 파일 로드 실패: ${file}`, fileErr.message);
        }
      }
    }

    manualCache = docs;
    console.log(`📚 메뉴얼 총 ${Object.keys(docs).length}개 로드 완료`);
  } catch (e) {
    console.error('메뉴얼 로드 실패:', e.message);
  }
}

function getManualsText() {
  if (!manualCache) return null;
  let text = '';
  for (const [name, content] of Object.entries(manualCache)) {
    text += `\n\n=== ${name} ===\n${content}`;
  }
  return text;
}

// ─── 회계 히스토리 파일 ────────────────────────────────────────────────────────
const ACCOUNTING_FILES = {
  '20':    { id: '1xHEI9J6a7ffJyQYDuyK-SOThLc-CjvBoIejUqn3aBR8', type: 'sheets' },  // 2020년 활동
  '22-2~': { id: '1At5GaEEv48L38Whacka1Dtv5ESgz4tFV',             type: 'xlsx'   },  // 회비 사용내역(22-2~)
  '25-2':  { id: '1x8JjScTIMtf4F8VmjPhJF6P0kzw6yIhbodKZ6ji5Xy8', type: 'sheets' },
  '26-1':  { id: '1qGqmCvod1_QDpiz3kQat6ZdfxM6dBL-_QGo4kOFClMg', type: 'sheets' },
};

// ─── Sheets 헬퍼 함수 ─────────────────────────────────────────────────────────
async function readSheet(spreadsheetId, range) {
  try {
    const sheets = await getSheetsClient();
    const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
    return res.data.values || [];
  } catch (e) {
    console.error('readSheet error:', e.message);
    return [];
  }
}

async function updateSheetRow(spreadsheetId, sheetName, rowIndex, values) {
  // rowIndex: 1-based (헤더=1, 첫 데이터=2)
  try {
    const sheets = await getSheetsClient();
    const colEnd = String.fromCharCode(64 + values.length); // A~E 등
    const range = `${sheetName}!A${rowIndex}:${colEnd}${rowIndex}`;
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [values] },
    });
    return true;
  } catch (e) {
    console.error('updateSheetRow error:', e.message);
    return false;
  }
}

async function appendSheet(spreadsheetId, range, values) {
  try {
    const sheets = await getSheetsClient();
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [values] },
    });
    return true;
  } catch (e) {
    console.error('appendSheet error:', e.message);
    return false;
  }
}

// ─── 이메일 알림 ──────────────────────────────────────────────────────────────
async function sendEmailNotification(subject, text) {
  if (!process.env.GMAIL_USER || !process.env.GMAIL_APP_PASSWORD) return;
  try {
    const transporter = nodemailer.createTransport({
      host: 'smtp.gmail.com',
      port: 465,
      secure: true,
      auth: { user: process.env.GMAIL_USER, pass: process.env.GMAIL_APP_PASSWORD },
    });
    await transporter.sendMail({
      from: process.env.GMAIL_USER,
      to: process.env.ADMIN_EMAIL || process.env.GMAIL_USER,
      subject,
      text,
    });
  } catch (e) {
    console.error('email error:', e.message);
  }
}

// ─── xlsx 파일 읽기 (Drive download + SheetJS 파싱) ──────────────────────────
async function readXlsxFromDrive(fileId) {
  try {
    const auth = await getAuthClient();
    const drive = google.drive({ version: 'v3', auth });
    const res = await drive.files.get(
      { fileId, alt: 'media' },
      { responseType: 'arraybuffer' }
    );
    const workbook = XLSX.read(Buffer.from(res.data), { type: 'buffer' });
    const result = {};
    for (const sheetName of workbook.SheetNames) {
      const ws = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      const nonEmpty = rows.filter(r => r.some(c => c !== ''));
      if (nonEmpty.length > 1) result[sheetName] = nonEmpty;
    }
    return result;
  } catch (e) {
    console.error('readXlsxFromDrive error:', e.message);
    return null;
  }
}

// ─── 회계 히스토리 조회 ───────────────────────────────────────────────────────
async function getAccountingHistory(period) {
  try {
    const sheets = await getSheetsClient();
    const results = {};

    const targets = period
      ? Object.entries(ACCOUNTING_FILES).filter(([k]) => k === period)
      : Object.entries(ACCOUNTING_FILES);

    for (const [term, file] of targets) {
      let termData;
      if (file.type === 'xlsx') {
        termData = await readXlsxFromDrive(file.id);
      } else {
        // Google Sheets 형식
        const meta = await sheets.spreadsheets.get({ spreadsheetId: file.id });
        const sheetNames = meta.data.sheets.map(s => s.properties.title);
        termData = {};
        for (const sheetName of sheetNames) {
          const res = await sheets.spreadsheets.values.get({
            spreadsheetId: file.id,
            range: `${sheetName}!A:G`,  // 실제 거래 열만 (H열 이후 계획표 제외)
          });
          const rows = res.data.values || [];
          if (rows.length > 1) termData[sheetName] = rows;
        }
      }
      if (termData) results[term] = termData;
    }
    return results;
  } catch (e) {
    console.error('getAccountingHistory error:', e.message);
    return null;
  }
}

// 파일명 매핑
const ACCOUNTING_FILE_NAMES = {
  '20':    '서울대 여자 라크로스 동아리 회계 내역',
  '22-2~': '회비 사용내역(22-2~).xlsx',
  '25-2':  '25-2 회계내역',
  '26-1':  '26-1 회계내역',
};

// 회계 데이터를 GPT에 넘길 수 있는 텍스트로 변환 (출처 포함)
function accountingDataToText(data) {
  if (!data) return '회계 데이터를 불러올 수 없어요.';
  let text = '[읽은 파일 목록]\n';
  for (const [term, sheets] of Object.entries(data)) {
    const fileName = ACCOUNTING_FILE_NAMES[term] || term;
    const tabNames = Object.keys(sheets).join(', ');
    text += `- "${fileName}" (탭: ${tabNames || '전체'})\n`;
  }
  text += '\n[거래 데이터 — 컬럼: 거래일시/분류/거래금액/대상/내용/거래후잔액/비고]\n';
  text += '※ 이 데이터는 실제 거래 내역만 포함합니다. H열 이후의 계획표/요약표는 포함되지 않습니다.\n';
  for (const [term, sheets] of Object.entries(data)) {
    const fileName = ACCOUNTING_FILE_NAMES[term] || term;
    text += `\n=== ${fileName} ===\n`;
    for (const [sheetName, rows] of Object.entries(sheets)) {
      text += `\n[탭: ${sheetName}]\n`;
      // 헤더 행, 예시 행(ex로 시작), 완전히 빈 행 제외
      const dataRows = rows.filter((r, i) => {
        if (i === 0) return false; // 헤더
        const first = (r[0] || '').toString().trim();
        if (first.startsWith('ex)') || first.startsWith('ex ')) return false; // 예시행
        if (!first) return false; // 빈 행
        return true;
      });
      const limited = dataRows.slice(0, 200);
      text += limited.map(r => r.slice(0, 7).join('\t')).join('\n'); // A:G만
      if (dataRows.length > 200) text += `\n... (총 ${dataRows.length}행)`;
      text += '\n';
    }
  }
  return text;
}

// ─── 공개 데이터 조회 함수들 ──────────────────────────────────────────────────
async function getPublicInfo(type) {
  switch (type) {
    case 'member_count': {
      const rows = await readSheet(SPREADSHEET_IDS.members, '부원 정보!A:A');
      return `현재 부원 수는 총 ${Math.max(0, rows.length - 1)}명이에요! 🥍`;
    }
    case 'training_schedule': {
      const rows = await readSheet(SPREADSHEET_IDS.training, '훈련 내용!A:E');
      if (rows.length <= 1) return '등록된 훈련 일정이 없어요.';
      const recent = rows.slice(-5).reverse();
      const lines = recent.map(r => `• ${r[0]} - ${r[2] || '내용 없음'}`).join('\n');
      return `최근 훈련 일정이에요! 🏃‍♂️\n${lines}`;
    }
    case 'equipment': {
      const rows = await readSheet(SPREADSHEET_IDS.equipment, '장비대여!A:D');
      return `현재 장비 대여 현황을 불러왔어요! 스틱과 볼 대여는 저한테 말씀해 주세요 🥍`;
    }
    default:
      return null;
  }
}

// ─── Google Calendar ──────────────────────────────────────────────────────────
const CALENDAR_ID = process.env.GOOGLE_CALENDAR_ID || 'primary';

async function createCalendarEvent({ title, date, startTime, endTime, description }) {
  try {
    const auth = await getAuthClient();
    const calendar = google.calendar({ version: 'v3', auth });
    const start = startTime
      ? { dateTime: `${date}T${startTime}:00+09:00`, timeZone: 'Asia/Seoul' }
      : { date };
    const end = endTime
      ? { dateTime: `${date}T${endTime}:00+09:00`, timeZone: 'Asia/Seoul' }
      : { date };
    const event = await calendar.events.insert({
      calendarId: CALENDAR_ID,
      requestBody: { summary: title, description, start, end },
    });
    return { success: true, link: event.data.htmlLink };
  } catch (e) {
    console.error('캘린더 이벤트 생성 실패:', e.message);
    return { success: false, error: e.message };
  }
}

async function getUpcomingEvents(days = 7) {
  try {
    const auth = await getAuthClient();
    const calendar = google.calendar({ version: 'v3', auth });
    const now = new Date();
    const later = new Date(now.getTime() + days * 24 * 60 * 60 * 1000);
    const res = await calendar.events.list({
      calendarId: CALENDAR_ID,
      timeMin: now.toISOString(),
      timeMax: later.toISOString(),
      singleEvents: true,
      orderBy: 'startTime',
      maxResults: 20,
    });
    return res.data.items || [];
  } catch (e) {
    console.error('캘린더 조회 실패:', e.message);
    return [];
  }
}

// ─── 게스트 승인 토큰 저장소 (메모리) ─────────────────────────────────────────
const pendingGuests = {}; // token → guestData

// ─── 게스트 신청 처리 ─────────────────────────────────────────────────────────
async function registerGuest(data) {
  const { name, department, studentId, date } = data;
  if (!name || !department || !studentId || !date) {
    return { success: false, message: '이름, 학과, 학번, 참여일자를 모두 입력해주세요!' };
  }

  // 승인 토큰 생성
  const token = Math.random().toString(36).slice(2) + Date.now().toString(36);
  pendingGuests[token] = { name, department, studentId, date, createdAt: new Date().toISOString() };

  const baseUrl = process.env.BASE_URL || `http://localhost:${process.env.PORT || 3000}`;
  const approveUrl = `${baseUrl}/api/guest-action?token=${token}&action=approve`;
  const rejectUrl  = `${baseUrl}/api/guest-action?token=${token}&action=reject`;

  // 알림 에이전트가 판단 후 이메일 발송
  runNotificationAgent('guest_signup', { name, department, studentId, date, approveUrl, rejectUrl }).catch(() => {});

  return { success: true, message: `${name}님 게스트 신청이 접수됐어요! 🎉 운영진 승인 후 확정돼요. 잠시만 기다려주세요~` };
}

// ─── 훈련 기록 저장 ───────────────────────────────────────────────────────────
async function saveTraining(data) {
  const { date, participants, attendees, content } = data;
  if (!date || !participants || !content) {
    return { success: false, message: '훈련일자, 참여인원수, 훈련내용을 모두 입력해주세요!' };
  }
  const row = [date, participants, attendees || '', content, new Date().toLocaleString('ko-KR')];
  const ok = await appendSheet(SPREADSHEET_IDS.training, '훈련 내용!A:E', row);
  if (!ok) return { success: false, message: '저장 중 오류가 발생했어요.' };

  // 구글 캘린더 등록은 백그라운드로
  createCalendarEvent({
    title: `🥍 SHALAC 훈련`,
    date,
    startTime: '18:00',
    endTime: '20:00',
    description: `참여인원: ${participants}명${attendees ? '\n참여자: ' + attendees : ''}\n내용: ${content}`,
  }).catch(() => {});

  return { success: true, message: `훈련 기록 저장 완료! 💪 ${date} 훈련 기록됐어요.` };
}

async function updateTraining(date, fields) {
  // 날짜로 기존 행 찾기
  const rows = await readSheet(SPREADSHEET_IDS.training, '훈련 내용!A:E');
  if (!rows || rows.length <= 1) return { success: false, message: '훈련 기록이 없어요.' };

  // 날짜 정규화 (2026-06-03, 06/03, 6/3 등 모두 대응)
  const normalize = (d) => String(d).replace(/\./g, '-').trim();
  const targetDate = normalize(date);

  // 헤더(0번) 제외하고 검색 → rowIndex는 시트 기준 (헤더=1, 데이터 시작=2)
  let foundIdx = -1;
  let foundRow = null;
  for (let i = 1; i < rows.length; i++) {
    if (normalize(rows[i][0] || '').includes(targetDate) || targetDate.includes(normalize(rows[i][0] || ''))) {
      foundIdx = i + 1; // 1-based sheet row
      foundRow = [...rows[i]];
      break;
    }
  }

  if (foundIdx === -1) return { success: false, message: `${date} 날짜의 훈련 기록을 찾지 못했어요. 날짜 형식을 확인해주세요.` };

  // 전달된 필드만 덮어쓰기 (없는 필드는 기존값 유지)
  // 열 순서: [0]훈련일자, [1]참여인원수, [2]참여자, [3]훈련내용, [4]기록일시
  if (fields.participants !== undefined) foundRow[1] = fields.participants;
  if (fields.attendees    !== undefined) foundRow[2] = fields.attendees;
  if (fields.content      !== undefined) foundRow[3] = fields.content;
  foundRow[4] = new Date().toLocaleString('ko-KR'); // 기록일시 갱신

  const ok = await updateSheetRow(SPREADSHEET_IDS.training, '훈련 내용', foundIdx, foundRow);
  return ok
    ? { success: true, message: `✅ ${date} 훈련 기록 업데이트 완료!\n수정된 내용: ${Object.keys(fields).join(', ')}` }
    : { success: false, message: '업데이트 중 오류가 발생했어요.' };
}

async function saveTrainingBulk(records) {
  const saved = [], failed = [];
  for (const data of records) {
    const { date, participants, attendees, content } = data;
    if (!date || !content) { failed.push(date || '날짜없음'); continue; }
    const row = [date, participants || '', attendees || '', content, new Date().toLocaleString('ko-KR')];
    const ok = await appendSheet(SPREADSHEET_IDS.training, '훈련 내용!A:E', row);
    if (ok) saved.push(date); else failed.push(date);
  }
  return { saved, failed };
}

// ─── 예산 기록 저장 (26-1 회계내역 양식: 거래일시/분류/거래금액/대상/내용/거래 후 잔액/비고) ──
async function saveBudget(data) {
  const { date, category, amount, target, content, note } = data;
  if (!date || !amount || !content) {
    return { success: false, message: '거래일시, 거래금액, 내용은 필수예요!' };
  }

  // 직전 행 잔액 읽어서 자동 계산
  const prevRows = await readSheet(SPREADSHEET_IDS.budget, '시트1!A:G');
  let prevBalance = 0;
  for (let i = prevRows.length - 1; i >= 1; i--) {
    const bal = prevRows[i][5];
    if (bal) {
      prevBalance = parseInt(bal.toString().replace(/,/g, ''), 10) || 0;
      break;
    }
  }
  const amtNum = parseInt(amount.toString().replace(/,/g, ''), 10) || 0;
  const newBalance = prevBalance + amtNum; // 입금 양수, 지출 음수로 입력

  const row = [
    date,
    category || '',
    amtNum,
    target || '',
    content,
    newBalance,
    note || '',
  ];

  const ok = await appendSheet(SPREADSHEET_IDS.budget, '시트1!A:G', row);
  if (!ok) return { success: false, message: '저장 중 오류가 발생했어요.' };

  // 알림 에이전트가 판단 후 이메일 발송 (백그라운드)
  runNotificationAgent('budget_record', { date, category, amount: amtNum, target, content, note, balance: newBalance }).catch(() => {});

  return { success: true, message: `예산 기록 완료! 💰 ${content} ${amtNum.toLocaleString()}원 → 잔액 ${newBalance.toLocaleString()}원` };
}

// ─── 장비 대여 기록 ───────────────────────────────────────────────────────────
async function saveEquipment(data) {
  const { borrower, equipment, borrowDate, returnDate } = data;
  if (!borrower || !equipment || !borrowDate) {
    return { success: false, message: '대여자, 장비명, 대여일을 입력해주세요!' };
  }
  const row = [borrower, equipment, borrowDate, returnDate || '미정', new Date().toLocaleString('ko-KR')];
  const ok = await appendSheet(SPREADSHEET_IDS.equipment, '장비 대여!A:E', row);
  return ok
    ? { success: true, message: `장비 대여 기록 완료! 🥍 ${borrower}님 ${equipment} 대여 기록됐어요.` }
    : { success: false, message: '저장 중 오류가 발생했어요.' };
}

// ─── 게스트 조회 (본인) ───────────────────────────────────────────────────────
async function checkGuestStatus(name, date) {
  const rows = await readSheet(SPREADSHEET_IDS.guests, '게스트 참여!A:F');
  if (rows.length <= 1) return '게스트 신청 내역이 없어요.';
  const found = rows.slice(1).filter(r => r[0] === name && r[3] === date);
  if (!found.length) return `${name}님의 ${date} 신청 내역을 찾을 수 없어요.`;
  const record = found[0];
  return `${name}님 게스트 신청 확인! ✅\n참여일: ${record[3]}\n입금 여부: ${record[4]}\n학과: ${record[1]} | 학번: ${record[2]}`;
}

// ─── 예산 계획 vs 실제 지출 비교 분석 ────────────────────────────────────────
async function analyzeBudgetVsPlan(spreadsheetId, category) {
  try {
    const sheets = await getSheetsClient();
    // 탭 목록 가져오기
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetNames = meta.data.sheets.map(s => s.properties.title);

    const result = { planned: [], actual: [], sources: [] };

    for (const sheetName of sheetNames) {
      // A:J 전체 읽기 (거래내역 + 계획표 모두 포함)
      const res = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `${sheetName}!A:J`,
      });
      const rows = res.data.values || [];
      result.sources.push(sheetName);

      // ① 계획 예산 추출: H~J열에서 카테고리 섹션 찾기
      let inSection = false;
      for (const row of rows) {
        const hCell = (row[7] || '').toString().trim();
        const iCell = (row[8] || '').toString().trim();
        const jCell = (row[9] || '').toString().trim();

        // 카테고리 헤더 감지 (예: "동소제", "훈련" 등)
        if (hCell === '' && iCell && !jCell) {
          inSection = iCell.includes(category) || category.includes(iCell);
        }
        // 금액|내용 헤더 행은 스킵
        if (iCell === '금액' && jCell === '내용') continue;
        // 총액 행
        if (iCell === '총액' && inSection) {
          const amt = parseAmount(jCell);
          if (amt > 0) result.planned.push({ label: '계획 총액', amount: amt });
          inSection = false;
          continue;
        }
        // 계획 항목 행
        if (inSection && iCell && jCell) {
          const amt = parseAmount(iCell);
          if (amt > 0) result.planned.push({ label: jCell, amount: amt });
        }
      }

      // ② 실제 지출 추출: A~G열 거래내역에서 카테고리 관련 항목 필터
      const keywords = category.replace(/\s/g, '');
      for (const row of rows) {
        const dateCell  = (row[0] || '').toString().trim();
        const classCell = (row[1] || '').toString().trim();
        const amtCell   = (row[2] || '').toString().trim();
        const targetCell= (row[3] || '').toString().trim();
        const descCell  = (row[4] || '').toString().trim();
        const balCell   = (row[5] || '').toString().trim();

        const combined = (classCell + descCell).replace(/\s/g, '');
        if (combined.includes(keywords)) {
          const amt = parseAmount(amtCell);
          if (amtCell && amtCell !== '거래금액') {
            result.actual.push({
              date: dateCell,
              label: descCell || classCell,
              target: targetCell,
              amount: amt,
              balance: balCell,
            });
          }
        }
      }
    }
    return result;
  } catch (e) {
    console.error('analyzeBudgetVsPlan error:', e.message);
    return null;
  }
}

// 금액 문자열 → 숫자 변환 ("50,000" → 50000, "-30,000" → -30000)
function parseAmount(str) {
  if (!str) return 0;
  const cleaned = str.toString().replace(/,/g, '').replace(/[₩\s]/g, '');
  return parseInt(cleaned, 10) || 0;
}

// ─── 훈련 출석 통계 ───────────────────────────────────────────────────────────
async function getTrainingAttendance(period, name) {
  const rows = await readSheet(SPREADSHEET_IDS.training, '훈련 내용!A:E');
  if (!rows || rows.length <= 1) return null;

  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  // 학기 구분: 1~6월 = 1학기, 7~12월 = 2학기
  const semesterStart = month <= 6
    ? new Date(year, 0, 1)   // 1월 1일
    : new Date(year, 6, 1);  // 7월 1일

  // 기간 필터 함수
  function inPeriod(dateStr) {
    if (!dateStr) return false;
    const d = new Date(String(dateStr).replace(/\./g, '-').replace(/(\d{4}-\d{2}-\d{2}).*/, '$1'));
    if (isNaN(d)) return false;
    if (period === 'semester') return d >= semesterStart;
    if (period === 'month')    return d.getFullYear() === year && d.getMonth() + 1 === month;
    if (period === 'year')     return d.getFullYear() === year;
    return true; // all
  }

  const dataRows = rows.slice(1).filter(r => r[0] && inPeriod(r[0]));

  // 참여자 파싱: C열(index 2)에 쉼표/공백 구분된 이름들
  const counter = {};
  const sessionCount = dataRows.length;

  for (const r of dataRows) {
    const attendeesRaw = String(r[2] || '');
    if (!attendeesRaw.trim()) continue;
    const names = attendeesRaw.split(/[,，、\s]+/).map(n => n.trim()).filter(Boolean);
    for (const n of names) {
      counter[n] = (counter[n] || 0) + 1;
    }
  }

  // 특정 인물 조회
  if (name) {
    const count = counter[name] || 0;
    return { type: 'personal', name, count, sessionCount, period };
  }

  // 랭킹
  const ranking = Object.entries(counter)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10)
    .map(([n, c], i) => ({ rank: i + 1, name: n, count: c }));

  return { type: 'ranking', ranking, sessionCount, period };
}

// ─── 기상청 단기예보 API ──────────────────────────────────────────────────────
// 서울대학교(관악구) 격자 좌표: nx=59, ny=124
const KMA_NX = 59, KMA_NY = 124;

function kmaFetch(url) {
  return new Promise((resolve, reject) => {
    const client = url.startsWith('https') ? https : http;
    client.get(url, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try { resolve(JSON.parse(data)); }
        catch { reject(new Error(`JSON 파싱 실패: ${data.slice(0, 200)}`)); }
      });
    }).on('error', reject);
  });
}

// 기상청 base_time 결정 (가장 최근 발표 시각)
function getKmaBaseTime() {
  const now = new Date();
  now.setMinutes(now.getMinutes() - 40); // 발표 후 40분 지연
  const h = now.getHours();
  const baseTimes = [2, 5, 8, 11, 14, 17, 20, 23];
  const base = baseTimes.filter(t => t <= h).pop() ?? 23;
  const baseDate = base === 23 && h < 23
    ? new Date(now.getTime() - 86400000) : now;
  const yyyy = baseDate.getFullYear();
  const mm = String(baseDate.getMonth() + 1).padStart(2, '0');
  const dd = String(baseDate.getDate()).padStart(2, '0');
  return { date: `${yyyy}${mm}${dd}`, time: String(base).padStart(2, '0') + '00' };
}

async function getWeatherForecast() {
  const apiKey = process.env.KMA_API_KEY;
  if (!apiKey) return null;

  const { date, time } = getKmaBaseTime();
  const url = `http://apis.data.go.kr/1360000/VilageFcstInfoService_2.0/getVilageFcst` +
    `?serviceKey=${encodeURIComponent(apiKey)}&numOfRows=1000&pageNo=1&dataType=JSON` +
    `&base_date=${date}&base_time=${time}&nx=${KMA_NX}&ny=${KMA_NY}`;

  try {
    const json = await kmaFetch(url);
    const items = json?.response?.body?.items?.item || [];

    // 날짜별로 그룹화
    const byDate = {};
    for (const item of items) {
      const d = item.fcstDate;
      if (!byDate[d]) byDate[d] = {};
      if (!byDate[d][item.fcstTime]) byDate[d][item.fcstTime] = {};
      byDate[d][item.fcstTime][item.category] = item.fcstValue;
    }

    // 각 날짜의 대표값 추출 (저녁 훈련 시간대 기준)
    const REPR_TIMES = ['1700', '2000', '2300']; // 저녁 시간대
    const PTY_LABEL = { '0': '맑음/흐림', '1': '비', '2': '비/눈', '3': '눈', '4': '소나기' };
    const SKY_LABEL = { '1': '맑음', '3': '구름많음', '4': '흐림' };

    const days = [];
    const today = new Date();

    for (const [fcstDate, times] of Object.entries(byDate).sort()) {
      const yyyy = fcstDate.slice(0,4), mm = fcstDate.slice(4,6), dd = fcstDate.slice(6,8);
      const dateObj = new Date(`${yyyy}-${mm}-${dd}`);
      const diffDay = Math.round((dateObj - new Date(today.toDateString())) / 86400000);
      if (diffDay < 0 || diffDay > 4) continue; // 오늘~4일 후까지

      const weekdays = ['일','월','화','수','목','금','토'];
      const weekday = weekdays[dateObj.getDay()];
      const label = diffDay === 0 ? '오늘' : diffDay === 1 ? '내일' : `${mm}/${dd}(${weekday})`;

      // 시간대별 개별 데이터
      const slots = REPR_TIMES.map(t => {
        const s = times[t] || {};
        const pty = s.PTY || '0';
        const wsd = parseFloat(s.WSD || '0');
        return {
          time: t.slice(0,2) + ':00',
          sky: SKY_LABEL[s.SKY] || '-',
          pty: PTY_LABEL[pty] || '-',
          hasRain: pty !== '0',
          tmp: s.TMP || null,
          wsd: wsd.toFixed(1),
          rain: (s.PCP && s.PCP !== '강수없음') ? s.PCP : '0',
          suitable: pty === '0' && wsd < 8,
        };
      });

      // 하루 전체 야외훈련 가능 여부 (모든 시간대 중 하나라도 비 오면 비추천)
      const suitable = slots.every(s => !s.hasRain) && slots.every(s => parseFloat(s.wsd) < 8);

      days.push({ label, date: `${mm}/${dd}`, weekday, slots, suitable });
    }
    return { days, baseDate: date, baseTime: time };
  } catch (e) {
    console.error('기상청 API 오류:', e.message);
    return null;
  }
}

function formatWeatherResult(result) {
  if (!result || !result.days || result.days.length === 0) return '날씨 정보를 가져오지 못했어요. KMA_API_KEY 설정을 확인해주세요.';

  const { days, baseDate, baseTime } = result;

  // 발표 시각 포맷 (예: "20260609", "1400" → "2026-06-09 14:00")
  const bd = baseDate || '';
  const bt = baseTime || '';
  const issuedAt = bd.length === 8
    ? `${bd.slice(0,4)}-${bd.slice(4,6)}-${bd.slice(6,8)} ${bt.slice(0,2)}:${bt.slice(2,4)}`
    : '';

  const lines = days.map(d => {
    const suitIcon = d.suitable ? '✅ 야외훈련 가능' : '❌ 야외훈련 비추천';
    const slotLines = d.slots.map(s => {
      const icon = s.hasRain ? '🌧️' : s.sky === '맑음' ? '☀️' : s.sky === '구름많음' ? '⛅' : '☁️';
      const rain = s.rain !== '0' ? ` 강수${s.rain}mm` : '';
      const tmp = s.tmp ? ` ${s.tmp}°C` : '';
      return `  · ${s.time} ${icon} ${s.sky}${s.hasRain ? '·'+s.pty : ''}${rain}${tmp} 바람 ${s.wsd}m/s`;
    }).join('\n');
    return `**${d.label}** ${suitIcon}\n${slotLines}`;
  });

  const suitableDays = days.filter(d => d.suitable).map(d => d.label);
  const summary = suitableDays.length > 0
    ? `\n\n야외 훈련 가능한 날: **${suitableDays.join(', ')}**`
    : '\n\n이번 주는 야외 훈련하기 어려운 날씨예요. 실내 훈련을 고려해보세요!';

  const source = `\n\n> 📡 출처: 기상청 단기예보 (발표: ${issuedAt} / 관측 지점: 서울 관악구 서울대)`;

  return `🌤️ **이번 주 날씨 & 훈련 가능 여부**\n\n${lines.join('\n')}${summary}${source}`;
}

// ─── 훈련 루틴 생성 ───────────────────────────────────────────────────────────
async function generateTrainingRoutine(participants, weatherContext) {
  const rows = await readSheet(SPREADSHEET_IDS.training, '훈련 내용!A:E');
  const history = rows.slice(1).slice(-10).map(r => `${r[0]}: ${r[3] || r[2]}`).join('\n'); // D열=훈련내용
  const weatherNote = weatherContext ? `\n오늘 날씨: ${weatherContext}` : '';
  const prompt = `당신은 라크로스 코치입니다. 아래는 최근 훈련 기록입니다:\n${history || '기록 없음'}${weatherNote}\n\n오늘 참여 인원: ${participants}명\n이를 바탕으로 오늘의 훈련 루틴을 상세히 짜주세요. 웜업, 기초 드릴, 팀 연습, 마무리 순서로 구성해주세요.`;
  return prompt;
}

// ─── Coordinator: 의도 분류 시스템 프롬프트 ──────────────────────────────────
const SYSTEM_PROMPT = `당신은 샤락이 🥍 — SHALAC 라크로스 동아리의 AI 어시스턴트예요!

[성격]
- 친근하고 활발해요. 라크로스라면 눈이 반짝반짝 빛나요!
- 구어체로 말해요. 이모지를 가끔 써요 (너무 많이는 NO).
- 라크로스 관련 질문엔 특히 열정적으로 답해요.
- 동아리원들의 든든한 조력자예요.

[역할]
다음 기능들을 도와줘요:
1. 게스트 신청 접수 (폼 안내 또는 직접 접수)
2. 훈련 내용 기록 (부원 인증 필요)
3. 예산 사용내역 기록 (부원 인증 필요)
4. 장비 대여 기록 (스틱, 볼 등)
5. 훈련 루틴 생성 (과거 기록 + 날씨 반영)
6. 날씨 예보 & 야외 훈련 가능 여부 안내 (서울대·관악구 기준)
6. 홍보글 작성 (인스타 캡션)
7. 동아리 소개제 준비 도움

[공개 정보 — 인증 없이 답변]
- 총 부원 수, 훈련 일정, 장비 현황
- 게스트 신청 방법 안내

[본인 조회 — 이름+날짜 입력받아 조회]
- 게스트 신청 확인, 입금 여부

[부원 인증 필요]
- 부원 상세 정보, 예산 내역, 훈련 기록 입력/수정

데이터가 없거나 기능이 준비 중이면 솔직하게 말하고, 운영진에게 연락하도록 안내해요.
JSON으로 응답하지 말고, 자연스러운 대화로 응답해요.

[한국라크로스협회 대회 기록]
대학리그 순위, 경기 결과, 팀/선수 스탯 관련 질문이 오면 get_kla_tournament 툴을 사용해요.
샤락(서울대&고려대 연합)의 경기 결과나 순위를 물어보면 반드시 최신 데이터를 조회해요.
알려진 대회 ID: 38 = 2025 대학리그 여자부, 37 = 2025 대학리그 남자부

[라크로스 규칙 안내 — 중요]
SHALAC은 여자 라크로스 동아리예요. 라크로스 규칙·반칙·경기 방식을 설명할 때는 반드시 여자 라크로스(Women's Lacrosse) 기준으로만 답해요. 룰북에 남자/여자 규정이 함께 있더라도 여자 기준 내용만 선택해서 답해요.

[회계 히스토리 — 절대 규칙]
예산, 지출, 회비, 잔액, 수입, 비교 등 회계 관련 질문이 오면 반드시 query_accounting_history 툴을 먼저 호출해요.
절대로 기억이나 추측으로 회계 수치를 답하지 않아요. 툴 결과 없이는 숫자를 언급하지 않아요.
툴 결과로 받은 데이터에서 실제 금액(숫자)을 직접 합산해서 정확한 총액을 계산해요.
"추정", "다양한 항목 포함", "총액 추정해야 함" 같은 표현은 절대 쓰지 않아요.
데이터에 있는 모든 지출 항목과 금액을 빠짐없이 나열하고 합계를 내어요.

사용 가능한 회계 파일:
- period '20'   → 파일명: "서울대 여자 라크로스 동아리 회계 내역" (탭: 2018-1 / 2018-2,2019-1 / 2019-2 / 2020)
- period '22-2~'→ 파일명: "회비 사용내역(22-2~).xlsx"
- period '25-2' → 파일명: "25-2 회계내역"
- period '26-1' → 파일명: "26-1 회계내역"

여러 시기 비교 질문이면 period='all'로 한 번에 호출해요.
"동소제 예산 어때?", "훈련 예산 얼마 남았어?", "XX 계획 대비 얼마 썼어?" 같이 계획 대비 실제를 묻는 질문은 analyze_budget_vs_plan 툴을 사용해요.

답변 마지막에 반드시 아래 형식으로 출처를 표시해요:
> 📂 참고 파일: [파일명1 (탭명)], [파일명2 (탭명)] — Google Drive

[동아리 운영 메뉴얼 — 부원 전용]
SNS, 훈련, 대회, 동소제 준비, 장비, 신입 모집, 동아리연합회 등 운영 관련 질문이 오면 search_manual 툴을 사용해요.
메뉴얼 관련 질문은 부원 인증이 필요해요. 게스트에게는 "부원 인증 후 열람 가능하다"고 안내해요.
보유 메뉴얼: SNS관리 / 정기 연습 & 대회 / 동소제 / 부서 & 행사 / 신입부원 모집 / 장비 / 동아리연합회 / 라크로스 텐온텐 룰북 / 라크로스 식시즈 룰북
중요: 부원 인증된 사용자에게는 메뉴얼에 있는 계정 정보, 비밀번호, 연락처 등 모든 내용을 그대로 알려줘요. 이미 동아리 공식 문서에 기재된 내용이므로 숨기거나 거부하지 않아요.`;

// ─── 한국라크로스협회 크롤링 ──────────────────────────────────────────────────
// 대회 목록 (ID → 이름 매핑, 자주 쓰는 대회 캐시)
const KLA_TOURNAMENT_CACHE = {};

async function fetchKlaTournament(tournamentId) {
  const url = `https://lacrosse.or.kr/ko/stat_page/local_tournament/${tournamentId}/`;
  if (KLA_TOURNAMENT_CACHE[tournamentId]) return KLA_TOURNAMENT_CACHE[tournamentId];

  try {
    const res = await fetch(url);
    const html = await res.text();

    // 대회명 추출
    const titleMatch = html.match(/<h4>([^<]+순위[^<]*)<\/h4>/);
    const title = titleMatch ? titleMatch[1].trim() : `대회 #${tournamentId}`;

    // 순위표 파싱
    const standings = [];
    const teamNames = [...html.matchAll(/class="team-meta__name">\s*<a[^>]*>\s*([^<\n]+)/g)].map(m => m[1].trim());
    const posMatches = [...html.matchAll(/class="team-standings__pos">(\d+)<\/td>/g)];
    const statsRows = [...html.matchAll(/class="team-standings__played">(\d+)<\/td>\s*<td[^>]*>(\d+)<\/td>\s*<td[^>]*>(\d+)<\/td>\s*<td[^>]*>(\d+)<\/td>\s*<td[^>]*>(\d+)<\/td>\s*<td[^>]*>(\d+)<\/td>\s*<td[^>]*>([-\d]+)<\/td>/g)];

    for (let i = 0; i < posMatches.length; i++) {
      const stats = statsRows[i];
      if (!stats) continue;
      standings.push({
        rank: posMatches[i][1],
        team: teamNames[i] || `팀 ${i+1}`,
        gp: stats[1], w: stats[2], l: stats[3], d: stats[4],
        gf: stats[5], ga: stats[6], gd: stats[7],
      });
    }

    // 경기 결과 파싱 (날짜)
    const dates = [...html.matchAll(/class="game-player-result__date">([^<]+)<\/td>/g)].map(m => m[1].trim());

    // 선수 스탯 (골/어시스트/GB)
    const playerStats = [];
    const playerRows = [...html.matchAll(/class="team-roster-table__name">([^<]+)<\/td>[\s\S]*?<td[^>]*>(\d+)<\/td>[\s\S]*?<td[^>]*>(\d+)<\/td>[\s\S]*?<td[^>]*>(\d+)<\/td>/g)];

    const result = { title, url, standings, dates, playerStats };
    KLA_TOURNAMENT_CACHE[tournamentId] = result;
    return result;
  } catch (e) {
    console.error('KLA 크롤링 실패:', e.message);
    return null;
  }
}

async function searchKlaTournament(query) {
  // HTML 전체를 GPT에 넘겨서 자연어로 파싱
  try {
    // 기본적으로 최근 대회 ID 목록 (38 = 2025 대학리그 여자부)
    const knownTournaments = [
      { id: 38, name: '2025 대학리그 여자부' },
      { id: 37, name: '2025 대학리그 남자부' },
      { id: 36, name: '2024 대학리그 여자부' },
      { id: 35, name: '2024 대학리그 남자부' },
    ];

    // 쿼리에서 관련 대회 찾기
    const matched = knownTournaments.find(t =>
      query.includes('여자') || query.includes('샤락') || query.includes('서울대')
        ? t.name.includes('여자')
        : t.name.includes(query) || query.includes(t.id.toString())
    ) || knownTournaments[0];

    const data = await fetchKlaTournament(matched.id);
    if (!data) return '한국라크로스협회 데이터를 불러오지 못했어요.';

    let result = `📊 **${data.title}** (출처: lacrosse.or.kr)\n\n`;

    if (data.standings.length > 0) {
      result += `**순위표**\n`;
      result += `| 순위 | 팀 | 경기 | 승 | 패 | 무 | 득점 | 실점 | 득실 |\n`;
      result += `|---|---|---|---|---|---|---|---|---|\n`;
      for (const s of data.standings) {
        result += `| ${s.rank} | ${s.team} | ${s.gp} | ${s.w} | ${s.l} | ${s.d} | ${s.gf} | ${s.ga} | ${s.gd} |\n`;
      }
    }

    if (data.dates.length > 0) {
      result += `\n**경기 일정**\n`;
      const uniqueDates = [...new Set(data.dates)];
      result += uniqueDates.map(d => `• ${d}`).join('\n');
    }

    result += `\n\n🔗 자세한 내용: ${data.url}`;
    return result;
  } catch (e) {
    return `대회 정보를 불러오지 못했어요: ${e.message}`;
  }
}

// ─── 알림 에이전트 ────────────────────────────────────────────────────────────
// 이벤트 발생 시 "이메일을 보낼지, 어떤 내용으로 보낼지" AI가 판단
async function runNotificationAgent(event, data) {
  try {
    const res = await openai.chat.completions.create({
      model: 'gpt-4o-mini',
      messages: [
        {
          role: 'system',
          content: `당신은 SHALAC 라크로스 동아리의 알림 에이전트예요.
이벤트와 데이터를 받아서 운영진에게 이메일 알림을 보내야 할지 판단하고,
보낸다면 제목과 내용을 작성해요.

판단 기준:
- guest_signup: 항상 알림 (운영진 승인 필요)
- budget_record: 항상 알림 (지출 확인 필요)
- budget_alert: 잔액이 30만원 미만이면 알림
- training_missing: 7일 이상 기록 없으면 알림
- training_record: 알림 불필요 (캘린더 등록으로 충분)
- exam_vote_reminder: 항상 알림 (시험기간 투표 진행 요청)
- recruit_reminder: 항상 알림 (신입부원 모집 준비 시작 안내, 학기 시작 3주 전)

반드시 아래 JSON 형식으로만 답하세요:
{"send": true/false, "subject": "제목", "body": "내용"}`,
        },
        {
          role: 'user',
          content: `이벤트: ${event}\n데이터: ${JSON.stringify(data, null, 2)}`,
        },
      ],
      temperature: 0,
      response_format: { type: 'json_object' },
    });

    const result = JSON.parse(res.choices[0].message.content);
    console.log(`📬 [알림 에이전트] ${event} → 발송: ${result.send}`);

    if (result.send && result.subject && result.body) {
      await sendEmailNotification(result.subject, result.body);
    }
    return result;
  } catch (e) {
    console.error('알림 에이전트 오류:', e.message);
    return { send: false };
  }
}

// ─── 멀티에이전트: 서브에이전트 시스템 프롬프트 ────────────────────────────────
const SUB_AGENT_PROMPTS = {
  operations: `당신은 샤락 운영 에이전트예요. SHALAC 라크로스 동아리의 훈련·예산·장비 기록을 전담해요.
친근한 구어체로 답해요. 훈련 기록, 예산 기록, 장비 대여 관련 작업을 정확하고 빠르게 처리해요.
부원 인증이 필요한 작업은 인증 여부를 반드시 확인해요.
라크로스 규칙 안내 시 여자 라크로스(Women's Lacrosse) 기준으로만 답해요.`,

  knowledge: `당신은 샤락 지식 에이전트예요. 라크로스 규칙, 동아리 운영 메뉴얼, 룰북을 전담해요.
친근한 구어체로 답해요. 정확한 규칙과 메뉴얼 내용을 근거와 함께 답해요.
라크로스 규칙은 반드시 여자 라크로스(Women's Lacrosse) 기준으로만 답해요.
메뉴얼 열람은 부원 인증이 필요해요.`,

  analytics: `당신은 샤락 분석 에이전트예요. 훈련 출석 통계, 예산 분석, 회계 히스토리를 전담해요.
친근한 구어체로 답해요. 데이터를 정확히 분석하고 인사이트를 제공해요.
예산/회계 질문엔 반드시 query_accounting_history 툴을 먼저 호출해요.
절대 추측으로 수치를 답하지 않아요. 툴 결과 없이 숫자를 언급하지 않아요.
답변 마지막에 반드시 출처를 표시해요: > 📂 참고 파일: [파일명] — Google Drive`,

  general: `당신은 샤락이 🥍 — SHALAC 라크로스 동아리의 AI 어시스턴트예요.
친근하고 활발해요. 구어체로 말해요. 이모지를 가끔 써요.
게스트 신청, 날씨, 동아리 소개, 홍보글 작성 등 일반적인 질문을 담당해요.`,
};

// ─── 오케스트레이터: 질문 유형 분류 ─────────────────────────────────────────
async function classifyQuery(userMessage) {
  try {
    const res = await openai.chat.completions.create({
      model: 'gpt-4o-mini',
      messages: [
        {
          role: 'system',
          content: `당신은 라크로스 동아리 AI 어시스턴트의 오케스트레이터예요.
사용자 메시지를 읽고 어떤 서브에이전트가 처리할지 분류하세요.

분류 기준:
- operations: 훈련 기록/수정, 예산 기록, 장비 대여 기록 등 데이터 입력/수정 작업
- knowledge: 라크로스 규칙, 룰북, 동아리 운영 메뉴얼 관련 질문
- analytics: 훈련 출석 통계, 예산 현황/분석, 회계 히스토리, 순위, 통계
- general: 날씨, 게스트 신청, 동아리 소개, 홍보글, 기타 일반 질문

반드시 operations / knowledge / analytics / general 중 하나만 답하세요.`,
        },
        { role: 'user', content: userMessage },
      ],
      temperature: 0,
      max_tokens: 10,
    });
    const agent = res.choices[0].message.content.trim().toLowerCase();
    if (['operations', 'knowledge', 'analytics', 'general'].includes(agent)) return agent;
    return 'general';
  } catch {
    return 'general';
  }
}

// ─── Tool definitions for OpenAI function calling ────────────────────────────
const TOOLS = [
  {
    type: 'function',
    function: {
      name: 'register_guest',
      description: '게스트 참여 신청을 접수하고 구글 시트에 저장합니다',
      parameters: {
        type: 'object',
        properties: {
          name: { type: 'string', description: '신청자 성명' },
          department: { type: 'string', description: '학과' },
          studentId: { type: 'string', description: '학번' },
          date: { type: 'string', description: '참여 희망 일자 (YYYY-MM-DD)' },
        },
        required: ['name', 'department', 'studentId', 'date'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'update_training',
      description: '기존 훈련 기록을 수정합니다. 이미 저장된 날짜의 참여자 이름, 인원수, 훈련내용 등 특정 필드만 업데이트할 때 사용합니다. 새 행을 추가하지 않고 기존 행을 덮어씁니다.',
      parameters: {
        type: 'object',
        properties: {
          date:         { type: 'string', description: '수정할 훈련 일자 (기존 기록과 일치해야 함, 예: 2026-06-03)' },
          participants: { type: 'number', description: '수정할 참여 인원수 (선택)' },
          attendees:    { type: 'string', description: '수정할 참여자 이름 목록 (쉼표 구분, 선택)' },
          content:      { type: 'string', description: '수정할 훈련 내용 (선택)' },
        },
        required: ['date'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'save_training_bulk',
      description: '여러 날짜의 훈련 기록을 한번에 저장합니다. 사용자가 표나 목록 형태로 여러 훈련 기록을 붙여넣으면 파싱해서 일괄 저장합니다 (부원 인증 필요).',
      parameters: {
        type: 'object',
        properties: {
          records: {
            type: 'array',
            description: '저장할 훈련 기록 배열',
            items: {
              type: 'object',
              properties: {
                date:         { type: 'string', description: '훈련 일자 (예: 2026-03-15)' },
                participants: { type: 'number', description: '참여 인원 수' },
                attendees:    { type: 'string', description: '참여자 이름 (쉼표 구분, 선택)' },
                content:      { type: 'string', description: '훈련 내용' },
              },
              required: ['date', 'content'],
            },
          },
        },
        required: ['records'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'save_training',
      description: '훈련 내용을 구글 시트에 저장합니다 (부원 인증 후 사용)',
      parameters: {
        type: 'object',
        properties: {
          date: { type: 'string', description: '훈련 일자 (예: 2026-06-09)' },
          participants: { type: 'number', description: '참여 인원 수 (숫자)' },
          attendees: { type: 'string', description: '참여자 이름 목록 (쉼표 구분, 예: 홍길동, 김철수, 이영희)' },
          content: { type: 'string', description: '훈련 내용 요약' },
        },
        required: ['date', 'participants', 'content'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'save_budget',
      description: '예산 사용 내역을 26-1 회계내역 시트에 저장합니다 (부원 인증 후 사용). 잔액은 자동 계산됩니다.',
      parameters: {
        type: 'object',
        properties: {
          date: { type: 'string', description: '거래일시 (예: 06/09)' },
          category: { type: 'string', description: '분류 (예: 회비 입금, 장비 구입, 코칭비, 동소제 등)' },
          amount: { type: 'number', description: '거래금액 (입금은 양수, 지출은 음수)' },
          target: { type: 'string', description: '대상 (거래 상대방 이름, 선택)' },
          content: { type: 'string', description: '내용 (거래 상세 내용)' },
          note: { type: 'string', description: '비고 (선택)' },
        },
        required: ['date', 'amount', 'content'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'save_equipment',
      description: '장비 대여 기록을 구글 시트에 저장합니다',
      parameters: {
        type: 'object',
        properties: {
          borrower: { type: 'string', description: '대여자 이름' },
          equipment: { type: 'string', description: '장비명 (스틱/볼 등)' },
          borrowDate: { type: 'string', description: '대여일' },
          returnDate: { type: 'string', description: '반납 예정일 (선택)' },
        },
        required: ['borrower', 'equipment', 'borrowDate'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'check_guest_status',
      description: '게스트 신청 현황 및 입금 여부를 조회합니다',
      parameters: {
        type: 'object',
        properties: {
          name: { type: 'string', description: '신청자 성명' },
          date: { type: 'string', description: '참여 신청 일자' },
        },
        required: ['name', 'date'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'get_public_info',
      description: '공개 정보를 조회합니다 (부원 수, 훈련 일정, 장비 현황)',
      parameters: {
        type: 'object',
        properties: {
          type: {
            type: 'string',
            enum: ['member_count', 'training_schedule', 'equipment'],
            description: '조회 유형',
          },
        },
        required: ['type'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'query_training_attendance',
      description: '훈련 출석 통계를 조회합니다. "훈련 제일 많이 나온 사람", "OOO 이번 학기 훈련 몇 번 나왔어?", "이번 달 출석 랭킹" 등에 사용합니다.',
      parameters: {
        type: 'object',
        properties: {
          period: {
            type: 'string',
            enum: ['semester', 'month', 'year', 'all'],
            description: '조회 기간. semester=이번학기, month=이번달, year=올해, all=전체',
          },
          name: {
            type: 'string',
            description: '특정 부원 이름 (개인 조회 시). 없으면 전체 랭킹 반환.',
          },
        },
        required: ['period'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'get_weather_forecast',
      description: '이번 주 날씨 예보를 조회하고 야외 훈련 가능 여부를 알려줍니다. "날씨 어때?", "이번 주 훈련 가능한 날은?", "비 오는 날은?" 등의 질문에 사용합니다.',
      parameters: {
        type: 'object',
        properties: {},
        required: [],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'get_calendar_events',
      description: '구글 캘린더에서 앞으로의 일정을 조회합니다. "다음 훈련 언제야?", "이번 주 일정 알려줘" 등의 질문에 사용합니다.',
      parameters: {
        type: 'object',
        properties: {
          days: { type: 'number', description: '조회할 일수 (기본 7일)', default: 7 },
        },
        required: [],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'get_kla_tournament',
      description: '한국라크로스협회 대회 순위, 경기결과, 팀 스탯을 조회합니다. "대학리그 순위", "서울대 경기 결과", "2025 여자부 순위" 등 대회 관련 질문에 사용합니다.',
      parameters: {
        type: 'object',
        properties: {
          query: { type: 'string', description: '조회할 대회명 또는 질문 (예: 2025 여자부, 서울대 경기결과)' },
          tournament_id: { type: 'number', description: '특정 대회 ID (알고 있는 경우, 예: 38)' },
        },
        required: ['query'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'generate_training_routine',
      description: '과거 훈련 기록과 오늘 날씨를 참고하여 오늘의 훈련 루틴을 생성합니다',
      parameters: {
        type: 'object',
        properties: {
          participants: { type: 'number', description: '오늘 참여 인원 수' },
          weather_context: { type: 'string', description: '오늘 날씨 요약 (선택, 예: 맑음 18°C 바람 2m/s)' },
        },
        required: ['participants'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'query_accounting_history',
      description: '과거 회계 내역을 조회하고 분석합니다. 특정 학기 지출, 카테고리별 분석, 동소제/훈련/장비 등 항목별 지출 조회에 사용합니다.',
      parameters: {
        type: 'object',
        properties: {
          period: {
            type: 'string',
            enum: ['20', '22-2~', '25-2', '26-1', 'all'],
            description: '조회할 시기. 20=2020년, 22-2~=22-2부터 누적, 25-2, 26-1, all=전체',
          },
          question: {
            type: 'string',
            description: '회계 데이터에서 찾고 싶은 내용 (예: 동소제 총 지출, 카테고리별 합계)',
          },
        },
        required: ['period', 'question'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'analyze_budget_vs_plan',
      description: '특정 카테고리(동소제, 훈련, 장비 등)의 계획 예산 대비 실제 지출을 비교 분석합니다. "동소제 예산 어때?", "훈련 예산 얼마 남았어?" 같은 질문에 사용합니다.',
      parameters: {
        type: 'object',
        properties: {
          period: {
            type: 'string',
            enum: ['25-2', '26-1'],
            description: '분석할 학기 (계획표가 있는 학기)',
          },
          category: {
            type: 'string',
            description: '분석할 예산 카테고리 (예: 동소제, 훈련, 장비)',
          },
        },
        required: ['period', 'category'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'search_manual',
      description: '동아리 운영 메뉴얼에서 관련 내용을 검색합니다. SNS, 훈련, 대회, 동소제, 장비, 신입부원 모집, 동아리연합회 등 운영 관련 질문에 사용합니다. 부원 인증 필요.',
      parameters: {
        type: 'object',
        properties: {
          query: { type: 'string', description: '찾고 싶은 내용 (예: 인스타 비밀번호, 대회 준비 방법, 장비 보관 위치)' },
          docName: {
            type: 'string',
            description: '특정 메뉴얼 지정 (선택). 없으면 전체 검색.',
            enum: ['SNS관리 메뉴얼', '정기 연습 & 대회 메뉴얼', '동소제 매뉴얼', '부서 & 행사 메뉴얼', '신입부원 모집 메뉴얼', '장비 메뉴얼', '동아리연합회 관련 메뉴얼', '라크로스 텐온텐 룰북', '라크로스 식시즈 룰북', '전체'],
          },
        },
        required: ['query'],
      },
    },
  },
];

// ─── Tool 실행기 ──────────────────────────────────────────────────────────────
async function executeTool(name, args, authLevel) {
  // 부원 인증이 필요한 기능 체크
  const memberOnlyTools = ['save_training', 'save_training_bulk', 'update_training', 'save_budget'];
  if (memberOnlyTools.includes(name) && authLevel !== 'member') {
    return '🔒 이 기능은 부원 인증이 필요해요! 비밀번호를 입력해주세요.';
  }

  switch (name) {
    case 'register_guest': {
      const result = await registerGuest(args);
      return result.message;
    }
    case 'update_training': {
      const { date, ...fields } = args;
      const result = await updateTraining(date, fields);
      return result.message;
    }
    case 'save_training_bulk': {
      if (!args.records?.length) return '저장할 기록이 없어요.';
      const { saved, failed } = await saveTrainingBulk(args.records);
      let msg = `✅ ${saved.length}건 저장 완료!\n저장된 날짜: ${saved.join(', ')}`;
      if (failed.length) msg += `\n\n⚠️ 실패: ${failed.join(', ')}`;
      return msg;
    }
    case 'save_training': {
      const result = await saveTraining(args);
      return result.message;
    }
    case 'save_budget': {
      const result = await saveBudget(args);
      return result.message;
    }
    case 'save_equipment': {
      const result = await saveEquipment(args);
      return result.message;
    }
    case 'check_guest_status': {
      return await checkGuestStatus(args.name, args.date);
    }
    case 'get_public_info': {
      return await getPublicInfo(args.type);
    }
    case 'query_training_attendance': {
      const result = await getTrainingAttendance(args.period, args.name);
      if (!result) return '훈련 기록이 없어요.';

      const periodLabel = { semester: '이번 학기', month: '이번 달', year: '올해', all: '전체' }[result.period] || result.period;

      if (result.type === 'personal') {
        return `[훈련 출석 조회]\n${periodLabel} 총 훈련 ${result.sessionCount}회 중 ${result.name}님은 **${result.count}회** 참여했어요.`;
      }

      if (!result.ranking.length) return `${periodLabel}에 참여자 데이터가 없어요. 훈련 기록에 참여자 이름이 입력됐는지 확인해주세요.`;

      const lines = result.ranking.map(r => `${r.rank}위 ${r.name} — ${r.count}회`).join('\n');
      return `[${periodLabel} 훈련 출석 랭킹] (총 ${result.sessionCount}회 훈련)\n\n${lines}`;
    }
    case 'get_weather_forecast': {
      const weatherResult = await getWeatherForecast();
      return formatWeatherResult(weatherResult);
    }
    case 'generate_training_routine': {
      return await generateTrainingRoutine(args.participants, args.weather_context);
    }
    case 'get_calendar_events': {
      const events = await getUpcomingEvents(args.days || 7);
      if (!events.length) return '앞으로 등록된 일정이 없어요.';
      const lines = events.map(e => {
        const start = e.start.dateTime
          ? new Date(e.start.dateTime).toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' })
          : e.start.date;
        return `• ${e.summary} — ${start}`;
      });
      return `📅 앞으로의 일정:\n${lines.join('\n')}`;
    }
    case 'get_kla_tournament': {
      if (args.tournament_id) {
        const data = await fetchKlaTournament(args.tournament_id);
        if (!data) return '대회 정보를 불러오지 못했어요.';
        let result = `📊 **${data.title}**\n\n순위표:\n`;
        for (const s of data.standings) {
          result += `${s.rank}위 ${s.team} — ${s.w}승 ${s.l}패 ${s.d}무 (득점 ${s.gf} / 실점 ${s.ga})\n`;
        }
        if (data.dates.length > 0) {
          result += `\n경기 일정: ${[...new Set(data.dates)].join(', ')}`;
        }
        result += `\n🔗 ${data.url}`;
        return result;
      }
      return await searchKlaTournament(args.query);
    }
    case 'query_accounting_history': {
      const targetPeriod = args.period === 'all' ? null : args.period;
      const data = await getAccountingHistory(targetPeriod);
      if (!data) return '회계 데이터를 불러오지 못했어요.';
      const text = accountingDataToText(data);
      return `[회계 데이터 로드 완료]\n질문: ${args.question}\n\n${text}`;
    }
    case 'analyze_budget_vs_plan': {
      const fileId = ACCOUNTING_FILES[args.period]?.id;
      if (!fileId) return `${args.period} 학기 데이터가 없어요.`;
      const fileName = ACCOUNTING_FILE_NAMES[args.period];
      const data = await analyzeBudgetVsPlan(fileId, args.category);
      if (!data) return '분석 중 오류가 발생했어요.';

      // 계획 총액 계산
      const planTotal = data.planned.find(p => p.label === '계획 총액')?.amount
        || data.planned.reduce((s, p) => s + p.amount, 0);
      const planItems = data.planned.filter(p => p.label !== '계획 총액');

      // 실제 지출 합계 (음수 = 지출)
      const actualTotal = data.actual.reduce((s, a) => s + Math.abs(a.amount), 0);
      const remaining = planTotal - actualTotal;
      const usedPct = planTotal > 0 ? ((actualTotal / planTotal) * 100).toFixed(1) : 0;

      let out = `[예산 분석: ${args.category} (${args.period}학기)]\n`;
      out += `참고 파일: "${fileName}" (탭: ${data.sources.join(', ')})\n\n`;

      out += `▸ 계획 예산\n`;
      if (planItems.length > 0) {
        planItems.forEach(p => { out += `  - ${p.label}: ${p.amount.toLocaleString()}원\n`; });
      }
      out += `  계획 총액: ${planTotal.toLocaleString()}원\n\n`;

      out += `▸ 실제 지출 내역\n`;
      if (data.actual.length === 0) {
        out += `  (관련 거래 없음)\n`;
      } else {
        data.actual.forEach(a => {
          out += `  - ${a.date} | ${a.label}${a.target ? ' (' + a.target + ')' : ''}: ${Math.abs(a.amount).toLocaleString()}원\n`;
        });
      }
      out += `  실제 지출 합계: ${actualTotal.toLocaleString()}원\n\n`;

      out += `▸ 비교 결과\n`;
      out += `  계획 대비 사용: ${usedPct}% (${actualTotal.toLocaleString()}원 / ${planTotal.toLocaleString()}원)\n`;
      out += remaining >= 0
        ? `  남은 예산: ${remaining.toLocaleString()}원\n`
        : `  초과 지출: ${Math.abs(remaining).toLocaleString()}원\n`;

      return out;
    }
    case 'search_manual': {
      if (authLevel !== 'member') {
        return '🔒 운영 메뉴얼은 부원 인증 후 열람 가능해요.';
      }
      if (!manualCache) {
        await loadManuals();
      }
      if (!manualCache) return '메뉴얼을 불러오지 못했어요. 잠시 후 다시 시도해주세요.';

      const { query, docName } = args;
      const targets = (!docName || docName === '전체')
        ? Object.entries(manualCache)
        : Object.entries(manualCache).filter(([name]) => name === docName);

      // 키워드 관련 섹션만 추출 (전체 텍스트가 너무 길 수 있으므로)
      let result = `[메뉴얼 검색: "${query}"]\n`;
      let found = false;
      for (const [name, content] of targets) {
        const lines = content.split('\n');
        const relevant = [];
        for (let i = 0; i < lines.length; i++) {
          if (lines[i].toLowerCase().includes(query.toLowerCase()) ||
              query.split(' ').some(kw => lines[i].includes(kw))) {
            // 앞뒤 2줄 포함
            const start = Math.max(0, i - 2);
            const end = Math.min(lines.length, i + 5);
            relevant.push(...lines.slice(start, end));
            relevant.push('---');
          }
        }
        if (relevant.length > 0) {
          result += `\n📄 ${name}\n${relevant.join('\n')}\n`;
          found = true;
        }
      }
      if (!found) {
        // 관련 섹션 못 찾으면 전체 내용 넘겨서 GPT가 판단
        result += '\n키워드 직접 매칭 없음. 전체 메뉴얼 참고:\n';
        for (const [name, content] of targets) {
          result += `\n📄 ${name}\n${content.slice(0, 2000)}\n`;
        }
      }
      result += `\n> 📂 출처: 동아리 운영 메뉴얼 — Google Drive`;
      return result;
    }
    default:
      return '알 수 없는 기능이에요.';
  }
}

// ─── 비밀번호 검증 ────────────────────────────────────────────────────────────
function verifyPassword(password) {
  return password === process.env.INTERNAL_PASSWORD;
}

// ─── 채팅 API ─────────────────────────────────────────────────────────────────
// ─── 쿼리 로깅 ───────────────────────────────────────────────────────────────
async function logQuery(query) {
  try {
    if (!SPREADSHEET_IDS.queries) return;
    const auth = await getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    const now = new Date().toISOString();
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_IDS.queries,
      range: '시트1!A:B',
      valueInputOption: 'RAW',
      requestBody: { values: [[now, query]] },
    });
  } catch (e) {
    // 로깅 실패는 무시
  }
}

// ─── 인기 검색어 API ──────────────────────────────────────────────────────────
app.get('/api/suggestions', async (req, res) => {
  const defaults = [
    '이번 주 날씨 어때?',
    '훈련 출석 순위 알려줘',
    '라크로스 반칙 규정 알려줘',
    '최근 훈련 내용 보여줘',
    '장비 대여 현황 알려줘',
    '예산 현황 알려줘',
  ];

  try {
    if (!SPREADSHEET_IDS.queries) return res.json({ suggestions: defaults });

    const auth = await getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_IDS.queries,
      range: '시트1!A:B',
    });

    const rows = result.data.values || [];
    if (rows.length < 5) return res.json({ suggestions: defaults });

    // 최근 7일 필터
    const weekAgo = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
    const recentQueries = rows
      .filter(row => row[0] && row[1] && new Date(row[0]) > weekAgo)
      .map(row => row[1].trim())
      .filter(q => q.length > 1 && q.length < 50);

    if (recentQueries.length < 5) return res.json({ suggestions: defaults });

    // 빈도 집계
    const freq = {};
    for (const q of recentQueries) {
      freq[q] = (freq[q] || 0) + 1;
    }
    const top6 = Object.entries(freq)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 6)
      .map(([q]) => q);

    // 6개 미만이면 defaults로 채움
    while (top6.length < 6) top6.push(defaults[top6.length]);

    res.json({ suggestions: top6 });
  } catch (e) {
    res.json({ suggestions: defaults });
  }
});

app.post('/api/chat', async (req, res) => {
  const { messages, password } = req.body;

  if (!messages || !Array.isArray(messages)) {
    return res.status(400).json({ error: 'messages 필드가 필요합니다.' });
  }

  // 사용자 마지막 메시지 로깅 (hidden_ 접두사 제외)
  const lastUserMsg = [...messages].reverse().find(m => m.role === 'user');
  if (lastUserMsg && !lastUserMsg.content?.startsWith?.('hidden_')) {
    logQuery(typeof lastUserMsg.content === 'string' ? lastUserMsg.content : JSON.stringify(lastUserMsg.content));
  }

  const authLevel = verifyPassword(password) ? 'member' : 'guest';

  // ── 오케스트레이터: 어떤 서브에이전트가 처리할지 분류 ──────────────────────
  const lastUserContent = lastUserMsg?.content;
  const userText = typeof lastUserContent === 'string' ? lastUserContent : '';
  const agentType = (!userText || userText.startsWith('hidden_'))
    ? 'general'
    : await classifyQuery(userText);

  console.log(`🤖 [오케스트레이터] "${userText.slice(0,40)}" → ${agentType} 에이전트`);

  // 서브에이전트 프롬프트 선택 (general은 기존 SYSTEM_PROMPT 사용)
  const subPrompt = agentType === 'general'
    ? SYSTEM_PROMPT
    : SUB_AGENT_PROMPTS[agentType] + '\n\n' + SYSTEM_PROMPT.split('[라크로스 규칙 안내')[1]
        ? SUB_AGENT_PROMPTS[agentType]
        : SYSTEM_PROMPT;

  const authNote = authLevel === 'member'
    ? '\n\n[현재 사용자 인증 상태: 부원 인증 완료 ✅ — 모든 기능 사용 가능. 인증을 요구하지 마세요.]'
    : '\n\n[현재 사용자 인증 상태: 게스트 — 부원 전용 기능(훈련 기록, 예산 기록) 요청 시 비밀번호 입력 안내.]';
  const dynamicPrompt = subPrompt + authNote;

  try {
    const response = await openai.chat.completions.create({
      model: 'gpt-4o-mini',
      messages: [{ role: 'system', content: dynamicPrompt }, ...messages],
      tools: TOOLS,
      tool_choice: 'auto',
      temperature: 0.8,
    });

    const assistantMessage = response.choices[0].message;

    // Tool 호출이 있을 경우 처리
    if (assistantMessage.tool_calls && assistantMessage.tool_calls.length > 0) {
      const toolResults = [];

      for (const toolCall of assistantMessage.tool_calls) {
        const toolName = toolCall.function.name;
        const toolArgs = JSON.parse(toolCall.function.arguments);
        const result = await executeTool(toolName, toolArgs, authLevel);

        toolResults.push({
          tool_call_id: toolCall.id,
          role: 'tool',
          content: result,
        });
      }

      // Tool 결과를 포함해 최종 응답 생성
      const finalResponse = await openai.chat.completions.create({
        model: 'gpt-4o-mini',
        messages: [
          { role: 'system', content: dynamicPrompt },
          ...messages,
          assistantMessage,
          ...toolResults,
        ],
        temperature: 0.8,
      });

      return res.json({
        reply: finalResponse.choices[0].message.content,
        authLevel,
        agentType,
      });
    }

    res.json({ reply: assistantMessage.content, authLevel, agentType });
  } catch (error) {
    console.error('OpenAI error:', error.message);
    res.status(500).json({ error: '죄송해요, 잠시 오류가 발생했어요. 다시 시도해주세요! 😅' });
  }
});

// ─── Drive 폴더 ID 매핑 ───────────────────────────────────────────────────────
// 서비스 계정 소유 드라이브 폴더 ID (서버 시작 시 자동 생성)
const DRIVE_FOLDERS = {
  root:      null,  // 샤락 에이전트 루트 폴더
  receipts:  null,  // 영수증
  training:  null,  // 훈련 사진
  general:   null,  // 일반 파일
};

// ─── 서비스 계정 드라이브 폴더 생성/조회 ─────────────────────────────────────
async function getOrCreateFolder(drive, name, parentId = null) {
  // 이미 존재하는 폴더 검색
  const q = parentId
    ? `name='${name}' and mimeType='application/vnd.google-apps.folder' and '${parentId}' in parents and trashed=false`
    : `name='${name}' and mimeType='application/vnd.google-apps.folder' and 'root' in parents and trashed=false`;

  const list = await drive.files.list({ q, fields: 'files(id,name)', pageSize: 1 });
  if (list.data.files.length > 0) return list.data.files[0].id;

  // 없으면 생성
  const meta = { name, mimeType: 'application/vnd.google-apps.folder' };
  if (parentId) meta.parents = [parentId];
  const created = await drive.files.create({ requestBody: meta, fields: 'id' });
  return created.data.id;
}

async function initDriveFolders() {
  try {
    const auth = await getAuthClient();
    const drive = google.drive({ version: 'v3', auth });
    const shareEmail = process.env.ADMIN_EMAIL || process.env.GMAIL_USER;

    // 폴더 구조 생성: 샤락 에이전트/ ├── 영수증/ ├── 훈련사진/ └── 일반파일/
    DRIVE_FOLDERS.root     = await getOrCreateFolder(drive, '샤락 에이전트');
    DRIVE_FOLDERS.receipts = await getOrCreateFolder(drive, '영수증',   DRIVE_FOLDERS.root);
    DRIVE_FOLDERS.training = await getOrCreateFolder(drive, '훈련사진', DRIVE_FOLDERS.root);
    DRIVE_FOLDERS.general  = await getOrCreateFolder(drive, '일반파일', DRIVE_FOLDERS.root);

    // 루트 폴더 공유 대상 목록
    const shareTargets = [
      shareEmail,
      process.env.DRIVE_SHARE_EMAIL,  // 추가 공유 계정
    ].filter(Boolean);

    for (const email of shareTargets) {
      try {
        await drive.permissions.create({
          fileId: DRIVE_FOLDERS.root,
          requestBody: { role: 'writer', type: 'user', emailAddress: email },
          sendNotificationEmail: false,
        });
        console.log(`📂 드라이브 폴더 공유 완료 → ${email}`);
      } catch (e) {
        if (!e.message.includes('already')) console.warn(`폴더 공유 경고 (${email}):`, e.message);
      }
    }

    console.log('📂 서비스 계정 드라이브 폴더 준비 완료');
  } catch (e) {
    console.error('initDriveFolders 오류:', e.message);
  }
}

// ─── Google Drive 파일 업로드 ─────────────────────────────────────────────────
async function uploadToDrive(fileBuffer, fileName, mimeType, folderId) {
  const auth = await getAuthClient();
  const drive = google.drive({ version: 'v3', auth });
  const { Readable } = require('stream');
  const stream = Readable.from(fileBuffer);
  const res = await drive.files.create({
    requestBody: { name: fileName, parents: [folderId] },
    media: { mimeType, body: stream },
    fields: 'id, name, webViewLink',
  });
  return res.data;
}

// ─── 영수증 OCR (GPT-4o Vision) ──────────────────────────────────────────────
// ─── 훈련 사진 인원 카운팅 (GPT-4o Vision) ───────────────────────────────────
async function countPeopleInPhoto(imageBuffer, mimeType) {
  const base64 = imageBuffer.toString('base64');
  const response = await openai.chat.completions.create({
    model: 'gpt-4o-mini',
    messages: [{
      role: 'user',
      content: [
        {
          type: 'text',
          text: `이 사진에서 사람의 수를 세어줘. 아래 JSON 형식으로만 답해줘.
{
  "count": 숫자 (보이는 사람 수, 부분적으로 보이는 사람도 포함),
  "confidence": "high" | "medium" | "low",
  "note": "특이사항 (선택, 예: 일부 가려진 인원 있음)"
}
JSON만 출력하고 다른 말은 하지 마.`,
        },
        { type: 'image_url', image_url: { url: `data:${mimeType};base64,${base64}` } },
      ],
    }],
    max_tokens: 100,
  });
  try {
    const text = response.choices[0].message.content.trim();
    return JSON.parse(text.replace(/```json|```/g, '').trim());
  } catch {
    return null;
  }
}

async function ocrReceipt(imageBuffer, mimeType) {
  const base64 = imageBuffer.toString('base64');
  const response = await openai.chat.completions.create({
    model: 'gpt-4o-mini',
    messages: [{
      role: 'user',
      content: [
        {
          type: 'text',
          text: `이 영수증 이미지를 분석해서 아래 JSON 형식으로 정확히 추출해줘. 확인 안 되는 필드는 null로 해줘. 오늘 날짜 기준 연도는 ${new Date().getFullYear()}년이야.
{
  "date": "YYYY-MM-DD 형식 날짜 (영수증에 연도 없으면 올해 연도 사용)",
  "category": "분류 (예: 장비 구입, 식비, 대관료, 코칭비 등)",
  "amount": 숫자만 (음수로, 예: -15000),
  "target": "거래처/상호명",
  "content": "구매 내용 요약",
  "note": "기타 참고사항"
}
JSON만 출력하고 다른 말은 하지 마.`,
        },
        {
          type: 'image_url',
          image_url: { url: `data:${mimeType};base64,${base64}` },
        },
      ],
    }],
    max_tokens: 300,
  });

  try {
    const text = response.choices[0].message.content.trim();
    const cleaned = text.replace(/```json|```/g, '').trim();
    return JSON.parse(cleaned);
  } catch {
    return null;
  }
}

// ─── 파일 업로드 API ──────────────────────────────────────────────────────────
app.post('/api/upload', upload.single('file'), async (req, res) => {
  const { password, mode } = req.body; // mode: 'receipt' | 'file'
  const isMember = verifyPassword(password);

  if (!isMember) {
    return res.status(403).json({ error: '부원 인증이 필요해요.' });
  }
  if (!req.file) {
    return res.status(400).json({ error: '파일이 없어요.' });
  }

  const { buffer, originalname, mimetype } = req.file;
  const isImage = mimetype.startsWith('image/');
  const timestamp = new Date().toLocaleDateString('ko-KR').replace(/\. /g,'-').replace('.','');
  const savedName = `${timestamp}_${originalname}`;

  try {
    // ① 드라이브 업로드 시도 (실패해도 OCR은 진행)
    let driveFile = null;
    let driveMsg = '';
    try {
      const folderId = mode === 'receipt'  ? DRIVE_FOLDERS.receipts
                     : mode === 'training' ? DRIVE_FOLDERS.training
                     : DRIVE_FOLDERS.general;
      driveFile = await uploadToDrive(buffer, savedName, mimetype, folderId);
      driveMsg = `\n📁 드라이브에도 저장됐어요.`;
    } catch (driveErr) {
      console.warn('Drive 업로드 실패 (OCR은 계속):', driveErr.message);
      driveMsg = '';
    }

    // ② 훈련 사진 모드 + 이미지면 인원 카운팅
    if (mode === 'training' && isImage) {
      const result = await countPeopleInPhoto(buffer, mimetype);
      return res.json({
        type: 'training_photo',
        driveFile,
        count: result,
        message: result
          ? `📸 사진 분석 완료!${driveMsg}\n\n인식된 인원: **${result.count}명** (신뢰도: ${result.confidence}${result.note ? ' / ' + result.note : ''})\n\n훈련 기록에 사용할까요?`
          : '사진에서 인원을 파악하기 어려웠어요. 직접 인원수를 알려주세요.',
      });
    }

    // ③ 영수증 모드 + 이미지면 OCR 실행
    if (mode === 'receipt' && isImage) {
      const ocr = await ocrReceipt(buffer, mimetype);
      return res.json({
        type: 'receipt',
        driveFile,
        ocr,
        message: ocr
          ? `🧾 영수증 인식 완료!${driveMsg}\n\n📋 인식 결과:\n• 날짜: ${ocr.date || '?'}\n• 분류: ${ocr.category || '?'}\n• 금액: ${ocr.amount?.toLocaleString() || '?'}원\n• 거래처: ${ocr.target || '?'}\n• 내용: ${ocr.content || '?'}\n\n회계에 기록할까요?`
          : `이미지를 받았는데 영수증 내용 인식이 어려웠어요. 내용을 직접 알려주시면 기록해드릴게요!`,
      });
    }

    // ③ 일반 파일 업로드
    if (driveFile) {
      return res.json({
        type: 'file',
        driveFile,
        message: `📁 "${originalname}" 드라이브에 저장 완료!\n🔗 ${driveFile.webViewLink}`,
      });
    } else {
      return res.json({
        type: 'file',
        message: `📎 파일을 받았어요! 드라이브 저장은 Shared Drive 설정이 필요해요. 영수증이라면 OCR 모드로 다시 올려주세요.`,
      });
    }

  } catch (e) {
    console.error('upload error:', e.message);
    res.status(500).json({ error: `업로드 실패: ${e.message}` });
  }
});

// ─── 비밀번호 확인 API ────────────────────────────────────────────────────────
app.post('/api/verify-password', (req, res) => {
  const { password } = req.body;
  res.json({ valid: verifyPassword(password) });
});

// ─── 게스트 승인/거절 엔드포인트 ──────────────────────────────────────────────
app.get('/api/guest-action', async (req, res) => {
  const { token, action } = req.query;
  const guest = pendingGuests[token];
  if (!guest) {
    return res.send('<h2>❌ 유효하지 않거나 이미 처리된 신청이에요.</h2>');
  }

  const { name, department, studentId, date } = guest;
  delete pendingGuests[token];

  if (action === 'approve') {
    const row = [name, department, studentId, date, '미입금', new Date().toLocaleString('ko-KR')];
    await appendSheet(SPREADSHEET_IDS.guests, '게스트 참여!A:F', row);
    await sendEmailNotification(
      `[샤락] 게스트 승인 완료: ${name}`,
      `${name}님의 게스트 신청이 승인되었어요!\n참여일: ${date}\n입금 안내를 진행해주세요.`
    );
    return res.send(`<h2>✅ ${name}님 게스트 신청을 승인했어요!</h2><p>시트에 기록됐어요. 입금 안내를 진행해주세요.</p>`);
  } else {
    await sendEmailNotification(
      `[샤락] 게스트 거절: ${name}`,
      `${name}님의 게스트 신청이 거절되었어요.\n참여일: ${date}`
    );
    return res.send(`<h2>❌ ${name}님 게스트 신청을 거절했어요.</h2>`);
  }
});

// 알림 테스트용 (부원 인증 필요)
app.post('/api/test-alerts', async (req, res) => {
  const { password } = req.body;
  if (!verifyPassword(password)) return res.status(403).json({ error: '인증 필요' });
  console.log('🧪 알림 테스트 수동 실행...');
  await Promise.all([checkBudgetAlert(), checkTrainingAlert()]);
  res.json({ ok: true, message: '알림 체크 완료! 조건 충족 시 메일이 발송됐어요.' });
});

// 월간 리포트 즉시 발송 테스트
app.post('/api/test-monthly-report', async (req, res) => {
  const { password } = req.body;
  if (!verifyPassword(password)) return res.status(403).json({ error: '인증 필요' });
  console.log('🧪 월간 리포트 수동 실행...');
  await sendMonthlyReport();
  res.json({ ok: true, message: '월간 리포트 메일이 발송됐어요!' });
});

// ─── 서버 시작 ────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
// ─── 선제적 알림 (Proactive Notifications) ───────────────────────────────────
const cron = require('node-cron');

// 잔액 30만원 미만 경고 체크
async function checkBudgetAlert() {
  try {
    const rows = await readSheet(SPREADSHEET_IDS.budget, '시트1!A:G');
    if (!rows || rows.length <= 1) return;

    // 마지막 유효 행에서 '거래 후 잔액' (F열, index 5) 읽기
    const dataRows = rows.slice(1).filter(r => r[0] && r[2]); // 날짜 + 금액 있는 행
    if (!dataRows.length) return;

    const lastRow = dataRows[dataRows.length - 1];
    const balanceRaw = String(lastRow[5] || '').replace(/[^0-9\-]/g, '');
    const balance = parseInt(balanceRaw, 10);

    if (isNaN(balance)) return;

    const THRESHOLD = 300000;
    if (balance < THRESHOLD) {
      await runNotificationAgent('budget_alert', {
        balance,
        lastDate: lastRow[0] || '-',
        lastContent: lastRow[4] || '-',
        lastAmount: parseInt(String(lastRow[2]).replace(/[^0-9\-]/g,''), 10) || 0,
      });
      console.log(`💸 예산 경고 알림 에이전트 호출 (잔액: ${balance.toLocaleString()}원)`);
    }
  } catch (e) {
    console.error('checkBudgetAlert error:', e.message);
  }
}

// 훈련 기록 1주 미기록 체크
async function checkTrainingAlert() {
  try {
    const rows = await readSheet(SPREADSHEET_IDS.training, '훈련 내용!A:E');
    if (!rows || rows.length <= 1) {
      await runNotificationAgent('training_missing', { lastDate: null, diffDays: null });
      return;
    }

    const dataRows = rows.slice(1).filter(r => r[0]); // 날짜 있는 행
    if (!dataRows.length) return;

    const lastDateStr = dataRows[dataRows.length - 1][0]; // 마지막 훈련 날짜
    // YYYY-MM-DD, MM/DD, YYYY.MM.DD 등 다양한 형식 파싱 시도
    const parsed = new Date(lastDateStr.replace(/\./g, '-').replace(/(\d{1,2})[-\/](\d{1,2})$/, `${new Date().getFullYear()}-$1-$2`));

    if (isNaN(parsed.getTime())) return;

    const diffDays = Math.floor((Date.now() - parsed.getTime()) / (1000 * 60 * 60 * 24));

    if (diffDays >= 7) {
      await runNotificationAgent('training_missing', { lastDate: lastDateStr, diffDays });
      console.log(`🏃 훈련 무기록 알림 에이전트 호출 (마지막: ${lastDateStr}, ${diffDays}일 전)`);
    }
  } catch (e) {
    console.error('checkTrainingAlert error:', e.message);
  }
}

// 월간 리포트 생성 및 발송
async function sendMonthlyReport() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1; // 방금 끝난 달
  const monthLabel = `${year}년 ${month}월`;

  try {
    // ① 예산 데이터 (26-1 시트)
    const budgetRows = await readSheet(SPREADSHEET_IDS.budget, '시트1!A:G');
    const budgetData = (budgetRows || []).slice(1).filter(r => r[0] && r[2]);

    // 이번 달 거래만 필터
    const monthStr = `${year}-${String(month).padStart(2,'0')}`;
    const thisMonthBudget = budgetData.filter(r => String(r[0]).startsWith(monthStr));

    const income  = thisMonthBudget.filter(r => parseFloat(String(r[2]).replace(/[^0-9\-]/g,'')) > 0);
    const expense = thisMonthBudget.filter(r => parseFloat(String(r[2]).replace(/[^0-9\-]/g,'')) < 0);
    const totalIncome  = income.reduce((s,r)  => s + Math.abs(parseFloat(String(r[2]).replace(/[^0-9\-]/g,''))||0), 0);
    const totalExpense = expense.reduce((s,r) => s + Math.abs(parseFloat(String(r[2]).replace(/[^0-9\-]/g,''))||0), 0);

    // 마지막 잔액
    const lastBalance = budgetData.length > 0
      ? String(budgetData[budgetData.length - 1][5] || '').replace(/[^0-9\-]/g,'')
      : '알 수 없음';

    // 지출 항목 상위 3개
    const expenseByCategory = {};
    for (const r of expense) {
      const cat = r[1] || '기타';
      expenseByCategory[cat] = (expenseByCategory[cat] || 0) + Math.abs(parseFloat(String(r[2]).replace(/[^0-9\-]/g,''))||0);
    }
    const topExpenses = Object.entries(expenseByCategory)
      .sort((a,b) => b[1]-a[1]).slice(0,3)
      .map(([k,v]) => `  - ${k}: ${v.toLocaleString()}원`).join('\n') || '  (없음)';

    // ② 훈련 기록
    const trainingRows = await readSheet(SPREADSHEET_IDS.training, '훈련 내용!A:E');
    const trainingData = (trainingRows || []).slice(1).filter(r => r[0]);
    const thisMonthTraining = trainingData.filter(r => String(r[0]).startsWith(monthStr));
    const totalSessions = thisMonthTraining.length;
    const totalParticipants = thisMonthTraining.reduce((s,r) => s + (parseInt(r[1])||0), 0);
    const avgParticipants = totalSessions > 0 ? (totalParticipants / totalSessions).toFixed(1) : 0;

    // ③ 리포트 메일 작성
    const subject = `[샤락] 📊 ${monthLabel} 월간 리포트`;
    const body = `안녕하세요! 샤락 에이전트가 ${monthLabel} 활동을 정리했어요 🥍

━━━━━━━━━━━━━━━━━━━━━━━━
💰 예산 현황
━━━━━━━━━━━━━━━━━━━━━━━━
• 이번 달 입금:  ${totalIncome.toLocaleString()}원
• 이번 달 지출:  ${totalExpense.toLocaleString()}원
• 현재 잔액:     ${parseInt(lastBalance||0).toLocaleString()}원

이번 달 주요 지출:
${topExpenses}

━━━━━━━━━━━━━━━━━━━━━━━━
🏃 훈련 현황
━━━━━━━━━━━━━━━━━━━━━━━━
• 총 훈련 횟수:  ${totalSessions}회
• 평균 참여 인원: ${avgParticipants}명
• 총 참여 연인원: ${totalParticipants}명

━━━━━━━━━━━━━━━━━━━━━━━━
다음 달도 화이팅! 💪
샤락 에이전트 드림 🥍
`;

    await sendEmailNotification(subject, body);
    console.log(`📊 ${monthLabel} 월간 리포트 발송 완료`);
  } catch (e) {
    console.error('sendMonthlyReport error:', e.message);
  }
}

// ── 스케줄 등록 ──────────────────────────────────────────────────────────────
// 매일 오전 9시: 잔액/훈련 알림 체크
cron.schedule('0 9 * * *', async () => {
  console.log('🔔 선제적 알림 체크 시작...');
  await Promise.all([checkBudgetAlert(), checkTrainingAlert()]);
  console.log('✅ 알림 체크 완료');
}, { timezone: 'Asia/Seoul' });

// ─── 학기말 리포트 ────────────────────────────────────────────────────────────
async function sendSemesterReport(semesterLabel, endMonth) {
  const now = new Date();
  const year = now.getFullYear();
  const startMonth = endMonth === 6 ? 3 : 9; // 1학기: 3~6월, 2학기: 9~12월

  try {
    // 해당 학기 월 범위
    const months = [];
    for (let m = startMonth; m <= endMonth; m++) {
      months.push(`${year}-${String(m).padStart(2,'0')}`);
    }

    // ① 예산
    const budgetRows = await readSheet(SPREADSHEET_IDS.budget, '시트1!A:G');
    const budgetData = (budgetRows || []).slice(1).filter(r => r[0] && r[2]);
    const semBudget = budgetData.filter(r => months.some(m => String(r[0]).startsWith(m)));
    const income  = semBudget.filter(r => parseFloat(String(r[2]).replace(/[^0-9\-]/g,'')) > 0);
    const expense = semBudget.filter(r => parseFloat(String(r[2]).replace(/[^0-9\-]/g,'')) < 0);
    const totalIncome  = income.reduce((s,r) => s + Math.abs(parseFloat(String(r[2]).replace(/[^0-9\-]/g,''))||0), 0);
    const totalExpense = expense.reduce((s,r) => s + Math.abs(parseFloat(String(r[2]).replace(/[^0-9\-]/g,''))||0), 0);
    const lastBalance  = budgetData.length > 0 ? parseInt(String(budgetData[budgetData.length-1][5]||'').replace(/[^0-9\-]/g,''))||0 : 0;

    const expenseByCategory = {};
    for (const r of expense) {
      const cat = r[1] || '기타';
      expenseByCategory[cat] = (expenseByCategory[cat] || 0) + Math.abs(parseFloat(String(r[2]).replace(/[^0-9\-]/g,''))||0);
    }
    const topExpenses = Object.entries(expenseByCategory)
      .sort((a,b) => b[1]-a[1]).slice(0,5)
      .map(([k,v]) => `  - ${k}: ${v.toLocaleString()}원`).join('\n') || '  (없음)';

    // ② 훈련
    const trainingRows = await readSheet(SPREADSHEET_IDS.training, '훈련 내용!A:E');
    const trainingData = (trainingRows || []).slice(1).filter(r => r[0]);
    const semTraining = trainingData.filter(r => months.some(m => String(r[0]).startsWith(m)));
    const totalSessions = semTraining.length;
    const totalParticipants = semTraining.reduce((s,r) => s + (parseInt(r[1])||0), 0);
    const avgParticipants = totalSessions > 0 ? (totalParticipants / totalSessions).toFixed(1) : 0;

    // 출석 순위 (상위 5명)
    const attendeeCount = {};
    for (const r of semTraining) {
      const names = (r[2] || '').split(/[,\s]+/).filter(n => n.trim());
      for (const n of names) attendeeCount[n] = (attendeeCount[n] || 0) + 1;
    }
    const topAttendees = Object.entries(attendeeCount)
      .sort((a,b) => b[1]-a[1]).slice(0,5)
      .map(([n,c], i) => `  ${i+1}위. ${n} (${c}회)`).join('\n') || '  (없음)';

    const subject = `[샤락] 📋 ${year} ${semesterLabel} 학기말 리포트`;
    const body = `안녕하세요! 샤락 에이전트가 ${year} ${semesterLabel} 활동을 정리했어요 🥍

━━━━━━━━━━━━━━━━━━━━━━━━
💰 학기 예산 현황
━━━━━━━━━━━━━━━━━━━━━━━━
• 총 수입:   ${totalIncome.toLocaleString()}원
• 총 지출:   ${totalExpense.toLocaleString()}원
• 현재 잔액: ${lastBalance.toLocaleString()}원

주요 지출 항목:
${topExpenses}

━━━━━━━━━━━━━━━━━━━━━━━━
🏃 학기 훈련 현황
━━━━━━━━━━━━━━━━━━━━━━━━
• 총 훈련 횟수:   ${totalSessions}회
• 평균 참여 인원: ${avgParticipants}명
• 총 참여 연인원: ${totalParticipants}명

이번 학기 출석 상위 5명:
${topAttendees}

━━━━━━━━━━━━━━━━━━━━━━━━
${semesterLabel} 한 학기 정말 수고했어요! 💪
샤락 에이전트 드림 🥍
`;

    await sendEmailNotification(subject, body);
    console.log(`📋 ${year} ${semesterLabel} 학기말 리포트 발송 완료`);
  } catch (e) {
    console.error('sendSemesterReport error:', e.message);
  }
}

// 매월 말일 오후 11시: 월간 리포트
cron.schedule('0 23 28-31 * *', async () => {
  const now = new Date();
  const tomorrow = new Date(now);
  tomorrow.setDate(tomorrow.getDate() + 1);
  // 내일이 다음 달 1일이면 오늘이 말일
  if (tomorrow.getDate() === 1) {
    console.log('📊 월간 리포트 생성 중...');
    await sendMonthlyReport();
  }
}, { timezone: 'Asia/Seoul' });

// 시험기간 투표 알림 (4월 1주 월요일, 5월 3주 월요일, 9월 1주 월요일, 11월 3주 월요일 오전 9시)
// 크론: 매주 월요일 오전 9시에 실행 → 날짜 조건 체크
cron.schedule('0 9 * * 1', async () => {
  const now = new Date(new Date().toLocaleString('en-US', { timeZone: 'Asia/Seoul' }));
  const month = now.getMonth() + 1; // 1~12
  const day = now.getDate();
  const week = Math.ceil(day / 7); // 몇째 주인지

  const examTargets = [
    { month: 4,  week: 1, label: '1학기 중간고사' },
    { month: 5,  week: 3, label: '1학기 기말고사' },
    { month: 9,  week: 1, label: '2학기 중간고사' },
    { month: 11, week: 3, label: '2학기 기말고사' },
  ];

  const recruitTargets = [
    { month: 2, week: 1, label: '1학기 신입모집' },
    { month: 8, week: 1, label: '2학기 신입모집' },
  ];

  const matchedExam = examTargets.find(t => t.month === month && t.week === week);
  if (matchedExam) {
    console.log(`📅 시험기간 투표 알림 발송: ${matchedExam.label}`);
    await runNotificationAgent('exam_vote_reminder', { examName: matchedExam.label, month, week });
  }

  const matchedRecruit = recruitTargets.find(t => t.month === month && t.week === week);
  if (matchedRecruit) {
    console.log(`📢 신입모집 준비 알림 발송: ${matchedRecruit.label}`);
    await runNotificationAgent('recruit_reminder', { label: matchedRecruit.label, month, week });
  }
}, { timezone: 'Asia/Seoul' });

// 학기말 리포트 (6월 2주, 12월 2주 월요일 오전 9시)
cron.schedule('0 9 * * 1', async () => {
  const now = new Date(new Date().toLocaleString('en-US', { timeZone: 'Asia/Seoul' }));
  const month = now.getMonth() + 1;
  const day = now.getDate();
  const week = Math.ceil(day / 7);

  if ((month === 6 || month === 12) && week === 2) {
    const semesterLabel = month === 6 ? '1학기' : '2학기';
    console.log(`📋 학기말 리포트 생성 중: ${semesterLabel}`);
    await sendSemesterReport(semesterLabel, month);
  }
}, { timezone: 'Asia/Seoul' });

console.log('🔔 선제적 알림 스케줄 등록됨 (매일 오전 9시 / 매월 말일 오후 11시 / 시험기간 투표 알림 / 학기말 리포트)');

app.listen(PORT, async () => {
  console.log(`🥍 샤락 에이전트 서버 실행 중: http://localhost:${PORT}`);
  await Promise.all([
    loadManuals(),
    initDriveFolders(),
  ]);
});
