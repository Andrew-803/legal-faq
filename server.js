require('dotenv').config();


const express = require('express');
const { GoogleGenerativeAI } = require('@google/generative-ai');
const path = require('path');
const XLSX = require('xlsx');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const MODEL_NAME = 'gemini-3.1-flash-lite';

function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) return val.toISOString().slice(0, 10);
  if (typeof val === 'number') {
    // Excel 날짜 시리얼 → JS Date (1900-01-01 기준, 윤년 버그 보정)
    const d = new Date(Math.round((val - 25569) * 86400 * 1000));
    return d.toISOString().slice(0, 10);
  }
  const s = String(val).split(' ')[0];
  return s.length >= 8 ? s : '';
}

// 실제 검토 내용 없이 첨부 파일만 안내하는 경우 감지
const FILE_ONLY_PATTERNS = [
  /첨부\s*파일\s*(을|로|과|와|은|는)?\s*(확인|참조|참고|대신|대체)/,
  /첨부\s*(문서|자료)\s*(를|을)?\s*(확인|참조|참고)/,
  /문서\s*파일\s*(로|을|으로)?\s*(작성하여\s*)?첨부/,
  /검토\s*의견\s*(을|은)?\s*첨부\s*(드립니다|파일|문서|자료)/,
  /답변\s*(은|를)?\s*첨부\s*파일/,
  /첨부와\s*같이\s*(의견을|작성)/,
  /첨부\s*자료를\s*참조/,
  /댓글\s*(을\s*통하여|로\s*통해|을\s*통해)\s*검토\s*의견/,
];

function detectFileOnly(reviewContents) {
  return FILE_ONLY_PATTERNS.some(p => p.test(reviewContents));
}

// ── Excel 데이터 로드 ──────────────────────────────────────────────
let faqDatabase = [];

function loadExcel() {
  try {
    const filePath = process.env.EXCEL_PATH ||
      path.join(__dirname, 'data', '법률자문_20260618.xlsx');

    const workbook = XLSX.readFile(filePath);
    const sheetName = 't_feed(본문내용)';
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) throw new Error(`시트를 찾을 수 없습니다: ${sheetName}`);

    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', cellDates: true });

    // 본문(댓글 제외)이고 검토가 완료된 건만 사용
    faqDatabase = rows.filter(r =>
      r['REPLY_YN'] === 'N' &&
      r['FEED_STATE'] === 'REVIEW_COMPLETED' &&
      String(r['REVIEW_CONTENTS'] || '').trim().length > 10
    ).map((r, idx) => {
      const reviewContents = String(r['REVIEW_CONTENTS'] || '').trim();
      return {
        id: idx,
        title: String(r['TITLE'] || '').trim(),
        requesterName: String(r['USER_NAME'] || '').trim(),
        requesterDept: String(r['ORG_NAME'] || '').trim(),
        contractPartner: String(r['CONTRACT_PARTNER'] || '').trim(),
        requestContents: String(r['REQUEST_CONTENTS'] || '').trim(),
        reviewContents,
        requestDate: formatDate(r['REQUEST_DATE']),
        reviewCompleteDate: formatDate(r['REVIEW_COMPLETE_DATE']),
        isFileOnly: detectFileOnly(reviewContents),
      };
    });

    console.log(`Excel 로드 완료: ${faqDatabase.length}건`);
  } catch (err) {
    console.error('Excel 로드 실패:', err.message);
    faqDatabase = [];
  }
}

loadExcel();

// ── 키워드 사전 필터 ───────────────────────────────────────────────
function keywordFilter(query) {
  const keywords = query.trim().split(/\s+/).filter(k => k.length >= 1);
  if (keywords.length === 0) return faqDatabase.slice(0, 20);

  const scored = faqDatabase.map(rec => {
    const searchTarget = [
      rec.title, rec.requestContents, rec.reviewContents, rec.contractPartner
    ].join(' ').toLowerCase();

    const score = keywords.reduce((sum, kw) => {
      const re = new RegExp(kw.toLowerCase().replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
      const matches = (searchTarget.match(re) || []).length;
      return sum + matches;
    }, 0);

    return { rec, score };
  });

  const topRecords = scored
    .filter(x => x.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, 15)
    .map(x => x.rec);

  // 키워드 매칭 건수가 적으면 전체에서 랜덤 보충
  if (topRecords.length < 3) {
    const extras = faqDatabase
      .filter(r => !topRecords.includes(r))
      .slice(0, 10 - topRecords.length);
    return [...topRecords, ...extras];
  }

  return topRecords;
}

// ── Gemini 호출 ───────────────────────────────────────────────────
function getModel() {
  const rawKey = process.env.GEMINI_API_KEY || '';
  const apiKey = rawKey.replace(/^﻿/, '').trim();
  if (!apiKey) throw new Error('.env 파일에 GEMINI_API_KEY가 설정되지 않았습니다.');
  const genAI = new GoogleGenerativeAI(apiKey);
  return genAI.getGenerativeModel({ model: MODEL_NAME });
}

function buildPrompt(query, records) {
  const recordsText = records.map((r, i) => {
    const title = r.title || '(제목 없음)';
    const dept = r.requesterDept;
    const date = r.reviewCompleteDate || r.requestDate;
    const req = r.requestContents.slice(0, 400);
    const review = r.reviewContents.slice(0, 600);
    const fileOnlyTag = r.isFileOnly ? ' [⚠️첨부파일만 있음]' : '';
    return `[사례 ${i + 1}] ${title}\n요청부서: ${dept} | 완료일: ${date}\n요청내용: ${req}\n검토의견${fileOnlyTag}: ${review}`;
  }).join('\n\n---\n\n');

  return `당신은 씨젠의료재단 준법감사팀의 법률자문 FAQ 검색 시스템입니다.
아래는 실제 법률자문 사례입니다. 사용자의 검색어와 가장 관련성 높은 사례를 찾아 FAQ 형식으로 정리해 주세요.

[사용자 검색어]: ${query}

[법률자문 사례]:
${recordsText}

위 사례 중 검색어와 관련된 것을 1~5개 선택하여 반드시 아래 JSON 형식으로만 응답하세요. JSON 외 다른 텍스트 없이.

※ 중요 규칙: 검토의견에 "[⚠️첨부파일만 있음]" 태그가 붙은 사례는 실제 검토 내용이 첨부 파일에만 있어 세부 내용을 확인할 수 없습니다.
이런 사례가 검색어와 관련성이 높을 경우, answer 필드를 반드시 다음과 같이 작성하세요:
"이 사례의 상세 검토 의견은 첨부 파일에만 제공되어 내용을 확인할 수 없습니다. 유사한 사안에 대해서는 준법감사팀에 직접 문의하시기 바랍니다."

{
  "totalFound": 관련 사례 총 건수(숫자),
  "summary": "검색 결과 한 줄 요약",
  "faqs": [
    {
      "question": "핵심 질문을 한 문장으로",
      "answer": "검토의견을 바탕으로 한 명확한 답변 (3~6문장)",
      "category": "계약검토|임대차|개인정보|노동·징계|공정거래|의료법|상표·IP|기타",
      "requesterDept": "요청부서명",
      "date": "자문완료일",
      "sourceTitle": "해당 사례의 원본 자문 제목(TITLE 그대로)"
    }
  ]
}`;
}

// ── API 엔드포인트 ─────────────────────────────────────────────────
app.post('/api/search', async (req, res) => {
  const { query } = req.body;
  if (!query || !query.trim()) {
    return res.status(400).json({ error: '검색어를 입력해 주세요.' });
  }

  if (faqDatabase.length === 0) {
    return res.status(503).json({ error: 'FAQ 데이터베이스가 로드되지 않았습니다. 서버 로그를 확인하세요.' });
  }

  try {
    const candidates = keywordFilter(query);
    const model = getModel();
    const prompt = buildPrompt(query, candidates);

    const result = await model.generateContent(prompt);
    const text = result.response.text().trim();

    const cleaned = text.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '');
    const jsonMatch = cleaned.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      return res.status(422).json({ error: 'AI 응답 파싱 실패', raw: text });
    }

    const data = JSON.parse(jsonMatch[0]);
    res.json(data);
  } catch (err) {
    console.error('검색 오류:', err.message);
    res.status(500).json({ error: `검색 중 오류: ${err.message}` });
  }
});

app.get('/api/stats', (req, res) => {
  res.json({
    total: faqDatabase.length,
    model: MODEL_NAME,
    status: faqDatabase.length > 0 ? 'ready' : 'error',
  });
});

app.post('/api/reload', (req, res) => {
  loadExcel();
  res.json({ total: faqDatabase.length });
});

if (require.main === module) {
  const PORT = process.env.PORT || 3002;
  app.listen(PORT, () => {
    console.log(`법률자문 FAQ 검색 서버 실행 중 → http://localhost:${PORT}`);
    console.log(`사용 모델: ${MODEL_NAME} | 데이터: ${faqDatabase.length}건`);
  });
}

module.exports = app;
