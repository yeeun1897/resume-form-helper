/* 이력서·자기소개서 양식 도우미
   모든 처리는 이 브라우저 안에서 해요. 네트워크 요청, 저장소(localStorage 등) 사용 없음. */
(function () {
  'use strict';

  const $ = (s, r = document) => r.querySelector(s);
  const { toParts } = window.DocxGen;

  // ── DOM 도우미: 사용자 글은 항상 textContent로 넣어요 ──
  function h(tag, attrs, ...kids) {
    const el = document.createElement(tag);
    if (attrs) {
      for (const [k, v] of Object.entries(attrs)) {
        if (v == null || v === false) continue;
        if (k === 'class') el.className = v;
        else if (k === 'text') el.textContent = v;
        else if (k === 'style') Object.entries(v).forEach(([p, val]) => el.style.setProperty(p, val));
        else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
        else el.setAttribute(k, v === true ? '' : v);
      }
    }
    for (const kid of kids.flat(Infinity)) {
      if (kid == null || kid === false) continue;
      el.append(kid instanceof Node ? kid : String(kid));
    }
    return el;
  }

  const ph = (t) => ({ ph: t });
  const rows = (n, cells) => Array.from({ length: n }, () => cells);

  const ORG_LABEL = {
    large: '국내 대기업·중견기업', public: '공기업·공공기관', foreign: '외국계 기업',
    startup: '스타트업·IT', creative: '디자인·마케팅·콘텐츠', research: '연구소·대학원·학계'
  };
  const LANG_LABEL = { ko: '국문', en: '영문', both: '국문+영문' };
  const ACCENTS = { navy: '#1F3A5F', charcoal: '#374151', green: '#1E5B4C', burgundy: '#6E2C3A' };

  // ── 입력 읽기 ──
  const form = $('#form');
  const fv = (name) => form.elements.namedItem(name);
  function readState() {
    const years = parseInt(fv('years').value, 10);
    return {
      org: fv('org').value,
      company: fv('company').value.trim().slice(0, 60),
      industry: fv('industry').value.trim().slice(0, 60),
      role: fv('jobrole').value.trim().slice(0, 60),
      lang: fv('doclang').value || 'ko',
      career: fv('career').value || 'new',
      years: Number.isFinite(years) ? Math.min(40, Math.max(0, years)) : 0,
      tone: fv('tone').value || 'classic',
      accent: fv('accent').value || 'navy',
      photo: fv('photo').checked,
      questions: parseQuestions(fv('questions').value),
    };
  }

  function parseQuestions(text) {
    return String(text || '').split(/\r?\n/).map((l) => l.trim()).filter(Boolean).slice(0, 10).map((line) => {
      const m = line.match(/(\d[\d,]{1,5})\s*(자|bytes?|바이트|words?|단어)/i);
      let limit = null;
      if (m) {
        const n = m[1].replace(/,/g, '');
        const num = Number(n).toLocaleString('ko-KR');
        limit = /자/.test(m[2]) ? `${num}자` : /바이트|byte/i.test(m[2]) ? `${num}바이트` : `${num}단어`;
      }
      const q = line
        .replace(/^(?:\d+\s*[.)]|[-*•]|Q\s*\d*\s*[.:)]?)\s*/i, '')
        .replace(/\s*[([（]\s*(?:최대\s*)?\d[\d,]*\s*(?:자|bytes?|바이트|words?|단어)\s*(?:이내|내외)?\s*[)\]）]\s*/i, ' ')
        .trim();
      return { q: q || line, limit };
    });
  }

  // ── 양식 이름 ──
  const NAMES = {
    'kr-table': { ko: '국문 표 형식 이력서' },
    'kr-blind': { ko: '블라인드 입사지원서' },
    'ncs-exp': { ko: 'NCS 경험·경력기술서' },
    'en-resume': { en: '영문 Resume' },
    'career-desc': { ko: '경력기술서형 이력서', en: '영문 프로젝트형 Resume' },
    'creative-one': { ko: '한 장 포트폴리오형 이력서', en: '영문 원페이지 크리에이티브 Resume' },
    'academic-cv': { ko: '국문 연구 이력서(CV)', en: 'Academic CV' },
    'kr-questions': { ko: '문항형 자기소개서' },
    'ncs': { ko: 'NCS 문항형 자기소개서' },
    'en-cover': { en: 'Cover Letter' },
    'free-intro': { ko: '자유 형식 자기소개서', en: 'Personal Statement' },
    'free-intro-portfolio': { ko: '포트폴리오 소개서', en: 'Portfolio Statement' },
    'research-statement': { ko: '연구계획서', en: 'Research Statement' },
  };
  function docName(id, lang, s) {
    const key = id === 'free-intro' && s.org === 'creative' ? 'free-intro-portfolio' : id;
    const n = NAMES[key];
    return n[lang] || n.ko || n.en;
  }

  // ── 추천 규칙 ──
  function langOrder(s) {
    if (s.lang === 'ko') return ['ko'];
    if (s.lang === 'en') return ['en'];
    return s.org === 'foreign' || s.org === 'research' ? ['en', 'ko'] : ['ko', 'en'];
  }
  const R = (resume, cover, lang, why) => ({ resume, cover, lang, why });

  function primary(s, lang) {
    const ko = lang === 'ko';
    switch (s.org) {
      case 'large':
        return ko
          ? R('kr-table', 'kr-questions', 'ko', ['국내 대기업·중견기업 공채는 인적사항·학력·경력을 표로 정리한 이력서를 가장 많이 써요.', '자기소개서는 회사가 낸 문항과 글자 수에 맞춰 쓰는 문항형이 기본이에요.'])
          : R('en-resume', 'en-cover', 'en', ['해외영업·글로벌 직무는 영문 이력서를 따로 요구하는 경우가 많아요.', '자기소개서 대신 한 장짜리 Cover Letter로 지원 동기를 정리해요.']);
      case 'public':
        return ko
          ? R('kr-blind', 'ncs', 'ko', ['공공기관은 블라인드 채용이라 사진·학교명·생년월일·출신지·가족 관계를 쓰지 않는 입사지원서를 써요.', '자기소개서는 NCS 직업기초능력(의사소통·문제해결·직업윤리 등)에 맞춘 문항형이에요.'])
          : R('en-resume', 'en-cover', 'en', ['국제협력·해외사업 직무는 영문 이력서를 함께 내기도 해요.', '사진·나이 없이 경력과 성과 중심으로 써요.']);
      case 'foreign':
        return ko
          ? R('kr-table', 'kr-questions', 'ko', ['외국계 한국 법인은 국문 이력서와 자기소개서를 함께 받는 경우가 많아요.', '국문 양식이라도 성과는 숫자로 짧게 쓰는 게 좋아요.'])
          : R('en-resume', 'en-cover', 'en', ['외국계는 사진·나이·성별 없이 경력과 성과를 중심으로 쓴 1~2장짜리 영문 Resume가 표준이에요.', 'Cover Letter 한 장에 지원 동기와 강점을 담아요.']);
      case 'startup':
        return R('career-desc', 'free-intro', lang, ['스타트업·IT는 프로젝트별로 문제-행동-결과를 보여 주는 경력기술서형 이력서가 유리해요.', '자기소개서는 정해진 문항 없이 짧고 구체적인 자유 형식을 선호해요.']);
      case 'creative':
        return R('creative-one', 'free-intro', lang, ['디자인·마케팅 직무는 강점·툴·대표 작업을 한 장에 한눈에 보여 주는 이력서가 좋아요.', '포트폴리오 링크와 대표 작업 설명을 담은 소개서를 함께 내요.']);
      case 'research':
        return R('academic-cv', 'research-statement', lang, ['연구직·대학원은 논문·학회 발표·연구 경력을 차례로 정리한 Academic CV를 써요.', '연구 질문과 앞으로의 계획을 담은 연구계획서(Research Statement)를 함께 내요.']);
      default:
        return null;
    }
  }

  function alternatives(s, L) {
    const first = L[0];
    const exp = s.career === 'exp';
    const koTable = R('kr-table', 'kr-questions', 'ko', ['표 형식 이력서를 요구하는 대기업 공채에도 함께 쓸 수 있어요.', '문항형 자기소개서를 미리 준비해 두면 공채 시즌에 빨라요.']);
    switch (s.org) {
      case 'large':
        return [R('career-desc', first === 'ko' ? 'free-intro' : 'en-cover', first, exp
          ? ['경력직 지원은 회사·프로젝트별 성과를 자세히 적는 경력기술서를 함께 요구하는 경우가 많아요.', '자유 양식 접수나 헤드헌터 제출에도 쓰기 좋아요.']
          : ['자유 양식 접수나 수시채용·추천 채용에 쓰기 좋은 경험 중심 이력서예요.', '인턴·프로젝트·대외활동을 성과 위주로 보여 줄 수 있어요.'])];
      case 'public':
        return L.includes('ko')
          ? [R('ncs-exp', 'ncs', 'ko', ['NCS 채용은 입사지원서와 별도로 경험·경력기술서를 받는 기관이 많아요.', '직무기술서의 필요 지식·기술과 이어지는 경험만 골라 써요.'])]
          : [R('career-desc', 'en-cover', 'en', ['프로젝트·경력을 자세히 보여 줘야 할 때 쓰는 영문 이력서예요.'])];
      case 'foreign':
        return [R('career-desc', first === 'en' ? 'en-cover' : 'free-intro', first, ['경력이나 프로젝트가 많다면 프로젝트별로 문제-행동-결과를 정리하는 형식이 강점을 잘 보여 줘요.', '헤드헌터·추천 지원 때도 쓰기 좋아요.'])];
      case 'startup':
      case 'creative':
        return [first === 'ko' ? koTable : R('en-resume', 'en-cover', 'en', ['해외 서비스·글로벌 팀에 지원할 때 쓰는 표준 영문 Resume예요.', '사진 없이 성과 중심으로 1~2장에 담아요.'])];
      case 'research':
        return [first === 'ko'
          ? R('kr-table', 'kr-questions', 'ko', ['기업 연구소(R&D)나 공채에 지원할 때는 표 형식 이력서와 문항형 자기소개서가 필요해요.'])
          : R('en-resume', 'en-cover', 'en', ['기업 연구소(R&D)에 지원할 때는 2장 이내의 Resume가 더 잘 맞아요.', '논문은 대표 몇 편만 골라 성과 중심으로 써요.'])];
      default:
        return [];
    }
  }

  function recommend(s) {
    if (!s.org) return [];
    const L = langOrder(s);
    const list = [];
    const add = (r, isPrimary) => {
      if (!r) return;
      if (list.some((x) => x.resume === r.resume && x.cover === r.cover && x.lang === r.lang)) return;
      list.push({ ...r, isPrimary });
    };
    L.forEach((lang) => add(primary(s, lang), true));
    alternatives(s, L).forEach((r) => add(r, false));
    return list.slice(0, 3).map((r, i) => ({
      ...r,
      key: `${r.resume}|${r.cover}|${r.lang}`,
      badge: i === 0 ? '가장 추천' : r.isPrimary ? '함께 준비' : '대안',
      resumeName: docName(r.resume, r.lang, s),
      coverName: docName(r.cover, r.lang, s),
    }));
  }

  function tipsFor(s, rec) {
    const t = [];
    const en = rec.lang === 'en';
    const exp = s.career === 'exp';
    if (en) {
      t.push('영문 서류에는 사진·생년월일·성별·가족 관계를 넣지 않아요.');
      t.push('각 줄은 Led, Built, Increased처럼 동사로 시작하고 결과를 숫자로 써요.');
      t.push(exp ? '경력이 길어도 2장을 넘기지 않아요.' : '신입은 1장 안에 담는 게 좋아요.');
    } else if (exp) {
      t.push(`${s.years ? `경력 ${s.years}년` : '경력직'}은 최근 경력부터 거꾸로 쓰고, 성과는 매출·비율·기간 같은 숫자로 적어요.`);
    } else {
      t.push('신입은 학력·교육·대외활동을 위쪽에 두고 1~2장 안에 담아요.');
    }
    if (s.photo && rec.resume === 'kr-blind') t.push('블라인드 채용이라 사진 칸을 넣지 않았어요.');
    if (s.photo && !['kr-table', 'kr-blind', 'creative-one'].includes(rec.resume)) t.push(en
      ? '영문 이력서에는 사진을 넣지 않는 게 일반적이라 사진 칸을 뺐어요.'
      : '이 양식은 사진 없이 쓰는 게 일반적이라 사진 칸을 뺐어요.');
    if (s.questions.length) {
      if (rec.cover === 'kr-questions' || rec.cover === 'ncs') t.push(`입력한 문항 ${s.questions.length}개로 자기소개서 칸을 만들었어요.`);
      else t.push('입력한 문항은 문항형 자기소개서(국문 양식)에만 들어가요.');
    }
    if (rec.resume === 'creative-one') t.push('포트폴리오 링크는 이력서 첫 화면에서 바로 보이게 두세요.');
    if (rec.resume === 'academic-cv') t.push('논문은 최신순으로 쓰고, 본인 이름은 굵게 표시해요.');
    if (rec.resume === 'kr-blind' || rec.cover === 'ncs') t.push('지원하는 기관의 직무기술서를 먼저 받아서, 거기 나온 단어를 그대로 써 주세요.');
    return t;
  }

  // ── 양식(문서 구조) ──
  function appliedLine(s, lang) {
    if (!s.company && !s.role) return null;
    const ko = lang === 'ko';
    return [s.company || ph(ko ? '회사명' : 'Company'), ' · ', s.role || ph(ko ? '지원 직무' : 'Position'), ko ? ' 지원' : ''];
  }

  function krTable(s, blind) {
    const exp = s.career === 'exp';
    const blocks = [{ t: 'title', text: blind ? '입사지원서' : '이력서', sub: appliedLine(s, 'ko') }];
    if (blind) blocks.push({ t: 'note', text: '블라인드 채용: 사진, 학교명, 생년월일, 출신지, 가족 관계처럼 개인을 알 수 있는 정보는 쓰지 않아요.' });
    const kvRows = blind ? [
      ['성명', ph('이름'), '연락처', ph('010-0000-0000')],
      ['이메일', ph('name@example.com'), '지원 분야', s.role || ph('지원 직무')],
      ['주소', ph('시·구까지만 (예: 서울시 마포구)')],
    ] : [
      ['성명(국문)', ph('홍길동'), '성명(영문)', ph('HONG GILDONG')],
      ['생년월일', ph('YYYY.MM.DD'), '연락처', ph('010-0000-0000')],
      ['이메일', ph('name@example.com'), '지원 분야', s.role || ph('지원 직무')],
      ['주소', ph('도로명 주소')],
    ];
    blocks.push({ t: 'section', title: '인적사항', blocks: [{ t: 'kv', rows: kvRows, photo: !blind && s.photo }] });

    const edu = blind
      ? { t: 'section', title: '학력사항', blocks: [
          { t: 'note', text: '학교명은 쓰지 않아요. 학력을 묻지 않는 기관이면 이 칸은 지워요.' },
          { t: 'table', head: ['기간', '최종 학력', '전공', '구분'], widths: [0.26, 0.24, 0.3, 0.2],
            rows: rows(2, [ph('YYYY.MM ~ YYYY.MM'), ph('고졸/대졸/석사'), ph('전공 계열'), ph('졸업/졸업예정')]) }] }
      : { t: 'section', title: '학력사항', blocks: [
          { t: 'table', head: ['기간', '학교명', '전공', '학점', '구분'], widths: [0.24, 0.22, 0.22, 0.14, 0.18],
            rows: rows(3, [ph('YYYY.MM ~ YYYY.MM'), ph('○○대학교'), ph('○○학과'), ph('0.0 / 4.5'), ph('졸업/졸업예정')]) }] };
    const career = { t: 'section', title: exp ? '경력사항' : '경력·인턴 경험', blocks: [
      { t: 'table', head: ['근무 기간', '회사명', '부서·직위', '담당 업무'], widths: [0.24, 0.22, 0.2, 0.34],
        rows: rows(exp ? 3 : 2, [ph('YYYY.MM ~ YYYY.MM'), ph(blind ? '기관·회사명' : '회사명'), ph('부서 / 직위'), ph('맡은 일과 성과 한 줄')]) }] };
    const cert = { t: 'section', title: '자격·어학', blocks: [
      { t: 'table', head: ['구분', '자격·시험명', '등급·점수', '취득일', '발행 기관'], widths: [0.14, 0.28, 0.18, 0.18, 0.22],
        rows: rows(3, [ph('자격/어학'), ph('예: 컴퓨터활용능력 1급'), ph('등급/점수'), ph('YYYY.MM'), ph('발행 기관')]) }] };
    const activity = { t: 'section', title: '교육·대외활동', blocks: [
      { t: 'table', head: ['기간', '활동명', '기관', '주요 내용'], widths: [0.24, 0.24, 0.18, 0.34],
        rows: rows(exp ? 2 : 3, [ph('YYYY.MM ~ YYYY.MM'), ph('활동·교육 이름'), ph('주관 기관'), ph('역할과 결과 한 줄')]) }] };
    const award = { t: 'section', title: '수상 경력', blocks: [
      { t: 'table', head: ['수상일', '수상명', '수여 기관', '내용'], widths: [0.16, 0.28, 0.22, 0.34],
        rows: rows(2, [ph('YYYY.MM'), ph('수상 이름'), ph('수여 기관'), ph('무엇으로 받았는지')]) }] };
    const military = { t: 'section', title: '병역', blocks: [
      { t: 'kv', rows: [['병역 구분', ph('군필 / 미필 / 면제 / 해당 없음'), '복무 기간', ph('YYYY.MM ~ YYYY.MM')]] }] };

    blocks.push(...(exp ? [career, edu, cert, activity, award, military] : [edu, career, cert, activity, award, military]));
    blocks.push({ t: 'sign', lines: ['위에 적은 내용은 사실과 다름이 없습니다.', ph('20○○년 ○월 ○일'), ['지원자  ', ph('이름'), '  (서명)']] });
    return blocks;
  }

  function ncsExp(s) {
    return [
      { t: 'title', text: '경험·경력기술서', sub: appliedLine(s, 'ko') },
      { t: 'note', text: '직무기술서에 적힌 필요 지식·기술과 관련된 경험만 골라 써요. 학교명이나 가족 정보처럼 개인을 알 수 있는 내용은 빼요.' },
      { t: 'section', title: '경력 (보수를 받고 일한 경험)', blocks: [
        { t: 'table', head: ['기간', '기관·회사', '직무·역할', '주요 업무와 성과'], widths: [0.2, 0.2, 0.2, 0.4],
          rows: rows(2, [ph('YYYY.MM ~ YYYY.MM'), ph('기관·회사'), ph('직무 / 역할'), ph('한 일과 결과')]) }] },
      { t: 'section', title: '경험 (동아리, 프로젝트, 봉사 등 보수 없이 한 활동)', blocks: [
        { t: 'table', head: ['기간', '활동명', '역할', '주요 내용과 성과'], widths: [0.2, 0.2, 0.2, 0.4],
          rows: rows(3, [ph('YYYY.MM ~ YYYY.MM'), ph('활동 이름'), ph('나의 역할'), ph('한 일과 결과')]) }] },
      { t: 'section', title: '직무 관련 교육', blocks: [
        { t: 'table', head: ['교육명', '교육 기관', '기간', '주요 내용'], widths: [0.26, 0.2, 0.2, 0.34],
          rows: rows(2, [ph('교육 이름'), ph('기관'), ph('YYYY.MM'), ph('배운 내용')]) }] },
      { t: 'section', title: '상세 기술', blocks: [
        { t: 'question', no: 1, q: '위 경력·경험 중 지원 직무와 가장 관련 있는 것을 골라 구체적으로 써 주세요.', limit: '500자',
          sub: '소제목: 핵심을 한 줄로', guide: '언제, 어디서, 어떤 역할로 → 무엇을 했고 → 어떤 결과(숫자)를 냈는지 순서로 써요.' }] },
    ];
  }

  const DEFAULT_Q = {
    corp: [
      { q: '지원 동기와 입사 후 포부', limit: '700자', guide: '왜 이 회사, 이 직무인지 → 그동안 준비한 것 → 입사 후 3~5년 목표 순서로 써요.' },
      { q: '직무 관련 역량과 경험', limit: '1,000자', guide: '직무에 필요한 역량 1~2가지를 골라 상황 → 행동 → 결과(숫자) 순서로 써요.' },
      { q: '가장 도전적이었던 경험', limit: '700자', guide: '목표, 부딪힌 어려움, 내가 한 행동, 결과와 배운 점을 써요.' },
      { q: '협업하며 갈등을 해결한 경험', limit: '700자', guide: '의견이 달랐던 상황, 내가 한 설득·조율, 팀이 얻은 결과를 써요.' },
    ],
    ncs: [
      { q: '지원 동기 (기관의 설립 목적·주요 사업과 연결해서)', limit: '500자', guide: '기관의 미션이나 최근 사업 → 나의 경험·관심 → 기여할 수 있는 점 순서로 써요.' },
      { q: '직무 관련 경험 (직무기술서의 필요 지식·기술 기준)', limit: '700자', guide: '직무기술서의 단어를 그대로 쓰고, 그 지식·기술을 써 본 경험을 구체적으로 적어요.' },
      { q: '의사소통·문제해결 능력을 발휘한 경험', limit: '600자', guide: '문제 상황 → 원인 분석 → 내가 한 행동 → 결과 순서로 써요.' },
      { q: '원칙과 직업윤리를 지킨 경험', limit: '500자', guide: '쉬운 길과 원칙 사이에서 원칙을 지킨 사례와 그 결과를 써요.' },
      { q: '입사 후 포부', limit: '500자', guide: '입사 후 1년, 3년, 5년 안에 이루고 싶은 목표를 기관 사업과 연결해요.' },
    ],
  };

  function krQuestions(s, mode) {
    const blocks = [
      { t: 'title', text: '자기소개서', sub: appliedLine(s, 'ko') },
      { t: 'kv', rows: [['지원 회사', s.company || ph('회사명'), '지원 직무', s.role || ph('지원 직무')], ['성명', ph('이름')]] },
    ];
    if (mode === 'ncs') blocks.push({ t: 'note', text: '블라인드 채용: 학교명, 출신지, 가족 직업처럼 개인을 알 수 있는 내용은 쓰지 않아요.' });
    const qs = s.questions.length
      ? s.questions.map((c) => ({ q: c.q, limit: c.limit, guide: '질문의 핵심 단어를 첫 문장에 넣고, 경험은 상황 → 행동 → 결과(숫자) 순서로 써요.' }))
      : DEFAULT_Q[mode];
    qs.forEach((x, i) => blocks.push({ t: 'question', no: i + 1, q: x.q, limit: x.limit, sub: '소제목: 핵심을 한 줄로', guide: x.guide }));
    return blocks;
  }

  function enResume(s) {
    const exp = s.career === 'exp';
    const summary = { t: 'section', title: 'Summary', blocks: [{ t: 'para', text: ph(exp
      ? `${s.years || 'N'} years of experience in your field — 2–3 lines on core strengths and one headline achievement`
      : `2–3 lines: your degree, strongest skills, and the value you bring to ${s.company || 'the company'}`) }] };
    const work = { t: 'section', title: exp ? 'Professional Experience' : 'Internships & Projects', blocks: rows(exp ? 3 : 2, null).map(() => ({
      t: 'entry', title: ph(exp ? 'Job Title' : 'Role / Project Title'), right: ph('MMM YYYY – Present'), sub: ph('Company or Organization · City'),
      bullets: [ph('Action verb + what you did + measurable result (e.g., Increased sales by 25%)'), ph('Action verb + project you led + impact'), exp ? ph('Scope, tools, or team size you managed') : null].filter(Boolean),
    })) };
    const edu = { t: 'section', title: 'Education', blocks: [{
      t: 'entry', title: ph('University Name'), right: ph('MMM YYYY'), sub: ph('B.A. in Major · GPA 0.0/4.5'),
      bullets: [ph('Relevant coursework, thesis, honors, or exchange program')] }] };
    const skills = { t: 'section', title: 'Skills & Certifications', blocks: [{ t: 'lines', items: [
      ['Languages', ph('Korean (Native), English (Fluent · TOEIC 000)')],
      ['Technical', ph('e.g., Excel, SQL, Tableau')],
      ['Certifications', ph('Certification name (Year)')]] }] };
    const acts = { t: 'section', title: 'Activities & Awards', blocks: [{ t: 'bullets', items: [
      ph('Award or activity — your role and result (Year)'), ph('Award or activity — your role and result (Year)')] }] };
    return [
      { t: 'header', name: ph('YOUR NAME'), sub: s.role || ph('Target Job Title'),
        contacts: [ph('Seoul, Korea'), ph('+82 10-0000-0000'), ph('name@example.com'), ph('linkedin.com/in/your-id')] },
      ...(exp ? [summary, work, edu, skills] : [summary, edu, work, acts, skills]),
    ];
  }

  function enCover(s) {
    return [
      { t: 'header', name: ph('Your Name'), contacts: [ph('name@example.com'), ph('+82 10-0000-0000'), ph('linkedin.com/in/your-id')] },
      { t: 'para', text: ph('Month DD, YYYY') },
      { t: 'addr', lines: [ph('Hiring Manager Name'), s.company || ph('Company Name'), ph('Company Address')] },
      { t: 'para', text: ['Dear ', ph('Ms. Lastname / Hiring Team'), ','] },
      { t: 'para', text: ph(`Opening: name the position (${s.role || 'Job Title'}), where you found it, and one sentence on why you are a strong fit.`) },
      { t: 'para', text: ph('Body 1: your most relevant experience, with one concrete and measurable result.') },
      { t: 'para', text: ph(`Body 2: why ${s.company || 'this company'} — connect its products, values, or recent news to your goals.`) },
      { t: 'para', text: ph('Closing: restate your interest, say you would welcome an interview, and thank the reader.') },
      { t: 'para', text: 'Sincerely,' },
      { t: 'para', text: ph('Your Name') },
    ];
  }

  function careerDesc(s, lang) {
    const T = (ko, en) => (lang === 'ko' ? ko : en);
    const exp = s.career === 'exp';
    const head = { t: 'header', name: ph(T('이름', 'YOUR NAME')), sub: s.role || ph(T('지원 직무', 'Target Job Title')),
      contacts: [ph(T('010-0000-0000', '+82 10-0000-0000')), ph('name@example.com'), ph(T('GitHub · 포트폴리오 링크', 'GitHub · Portfolio URL'))] };
    const summary = { t: 'section', title: T('요약', 'Summary'), blocks: [{ t: 'para', text: ph(exp
      ? T(`경력 ${s.years || 'N'}년 · 핵심 역량 3가지 · 대표 성과 1가지를 2~3줄로`, `${s.years || 'N'} years of experience · 3 core strengths · 1 headline result, in 2–3 lines`)
      : T('어떤 문제를 잘 푸는 사람인지 + 대표 경험 하나를 2~3줄로', '2–3 lines: the problems you solve best + one standout project')) }] };
    const skills = { t: 'section', title: T('기술·역량', 'Skills'), blocks: [{ t: 'lines', items: [
      [T('주요 기술', 'Core'), ph(T('예: Python, SQL, React', 'e.g., Python, SQL, React'))],
      [T('도구', 'Tools'), ph(T('예: Git, Jira, Figma', 'e.g., Git, Jira, Figma'))],
      [T('자격·어학', 'Certs & Languages'), ph(T('자격증 / 어학 점수', 'Certifications / language scores'))]] }] };
    const career = { t: 'section', title: T('경력', 'Experience'), blocks: rows(2, null).map(() => ({
      t: 'entry', title: ph(T('회사명', 'Company Name')), right: ph(T('YYYY.MM ~ 현재', 'MMM YYYY – Present')),
      sub: ph(T('부서 · 직급 · 담당 업무', 'Team · Title · Scope')),
      bullets: [ph(T('맡은 일: 담당 서비스·업무 범위', 'Scope: the product or area you owned')), ph(T('성과: 숫자로 (예: 응답 속도 40% 개선)', 'Result: in numbers (e.g., cut latency by 40%)'))] })) };
    const projects = { t: 'section', title: T('프로젝트', 'Projects'), blocks: rows(exp ? 2 : 3, null).map(() => ({
      t: 'entry', title: ph(T('프로젝트명', 'Project Name')), right: ph(T('YYYY.MM ~ YYYY.MM', 'MMM YYYY – MMM YYYY')),
      sub: ph(T('소속 · 역할 · 참여 인원', 'Team · Role · Team size')),
      bullets: [ph(T('문제: 해결하려던 문제', 'Problem: what needed solving')), ph(T('행동: 내가 한 일과 쓴 기술', 'Action: what you did and the tools you used')), ph(T('결과: 숫자로 보여 주는 성과', 'Result: a measurable outcome'))] })) };
    const edu = { t: 'section', title: T('학력', 'Education'), blocks: [{ t: 'entry', title: ph(T('학교명', 'University Name')),
      right: ph(T('YYYY.MM 졸업', 'MMM YYYY')), sub: ph(T('전공 · 학위', 'Degree · Major')) }] };
    return [head, ...(exp ? [summary, skills, career, projects, edu] : [summary, skills, projects, edu])];
  }

  function creativeOne(s, lang) {
    const T = (ko, en) => (lang === 'ko' ? ko : en);
    const side = [
      s.photo ? { t: 'photo' } : null,
      { t: 'section', title: T('연락처', 'CONTACT'), blocks: [{ t: 'lines', stacked: true, items: [
        [T('전화', 'Phone'), ph(T('010-0000-0000', '+82 10-0000-0000'))], [T('이메일', 'Email'), ph('name@example.com')],
        [T('포트폴리오', 'Portfolio'), ph('URL')], ['SNS', ph(T('Behance · Instagram 등', 'Behance · Instagram'))]] }] },
      { t: 'section', title: T('도구', 'TOOLS'), blocks: [{ t: 'bullets', items: [ph('Figma'), ph('Photoshop · Illustrator'), ph(T('예: Premiere Pro, GA4', 'e.g., Premiere Pro, GA4'))] }] },
      { t: 'section', title: T('강점', 'STRENGTHS'), blocks: [{ t: 'bullets', items: [ph(T('강점 키워드 1', 'Strength 1')), ph(T('강점 키워드 2', 'Strength 2')), ph(T('강점 키워드 3', 'Strength 3'))] }] },
      { t: 'section', title: T('학력', 'EDUCATION'), blocks: [{ t: 'entry', title: ph(T('학교명', 'University')), sub: ph(T('전공 · 졸업 연도', 'Major · Year')) }] },
      { t: 'section', title: T('어학·자격', 'LANGUAGES'), blocks: [{ t: 'bullets', items: [ph(T('예: TOEIC 000', 'e.g., English — Fluent'))] }] },
    ].filter(Boolean);
    const main = [
      { t: 'header', name: ph(T('이름', 'YOUR NAME')), sub: s.role || ph(T('지원 직무 (예: 브랜드 디자이너)', 'Target Role (e.g., Brand Designer)')) },
      { t: 'section', title: T('소개', 'PROFILE'), blocks: [{ t: 'para', text: ph(T('나의 작업 스타일과 강점을 2~3줄로', '2–3 lines on your style and strengths')) }] },
      { t: 'section', title: T('경력·프로젝트', 'EXPERIENCE & PROJECTS'), blocks: rows(3, null).map(() => ({
        t: 'entry', title: ph(T('프로젝트·회사명', 'Project or Company')), right: ph('YYYY.MM'), sub: ph(T('역할 · 사용 툴', 'Role · Tools')),
        bullets: [ph(T('무엇을 만들었는지', 'What you made')), ph(T('결과: 조회수, 전환율, 수상 등 숫자로', 'Result: views, conversion, awards'))] })) },
      { t: 'section', title: T('대표 작업', 'SELECTED WORK'), blocks: [{ t: 'lines', items: [1, 2, 3].map((n) => [ph(T(`작업명 ${n}`, `Work ${n}`)), ph(T('링크 – 한 줄 설명', 'URL – one-line description'))]) }] },
      { t: 'section', title: T('수상·전시', 'AWARDS & EXHIBITIONS'), blocks: [{ t: 'bullets', items: rows(2, ph(T('수상명 / 전시명 (연도)', 'Award or exhibition (Year)'))) }] },
    ];
    return [{ t: 'twocol', side, main }];
  }

  function academicCv(s, lang) {
    const T = (ko, en) => (lang === 'ko' ? ko : en);
    return [
      { t: 'header', name: ph(T('이름', 'YOUR NAME')), sub: ph(T('소속 · 직위 (예: ○○대학교 박사과정)', 'Affiliation · Position (e.g., Ph.D. Candidate)')),
        contacts: [ph('name@university.example'), ph(T('ORCID · Google Scholar 링크', 'ORCID · Google Scholar')), ph(T('연구실 홈페이지', 'Lab or personal website'))] },
      { t: 'section', title: T('연구 관심 분야', 'Research Interests'), blocks: [{ t: 'para', text: ph(T('핵심 키워드 3~5개와 한두 줄 설명', '3–5 keywords with a one-line description')) }] },
      { t: 'section', title: T('학력', 'Education'), blocks: rows(2, null).map(() => ({
        t: 'entry', title: ph(T('학위 · 전공 (예: 공학석사, 기계공학)', 'Degree, Field (e.g., M.S. in Mechanical Engineering)')), right: ph('YYYY'),
        sub: ph(T('학교명', 'University Name')), bullets: [ph(T('학위논문 제목 · 지도교수', 'Thesis title · Advisor'))] })) },
      { t: 'section', title: T('논문·출판', 'Publications'), blocks: [{ t: 'bullets', numbered: true, items: rows(3,
        ph(T('저자, 저자 (연도). 논문 제목. 학술지명, 권(호), 쪽. — 본인 이름은 굵게', 'Author, A., & Author, B. (Year). Title. Journal, Vol(Issue), pp. — bold your name'))) }] },
      { t: 'section', title: T('학회 발표', 'Conference Presentations'), blocks: [{ t: 'bullets', items: rows(2,
        ph(T('발표자 (연도). 발표 제목. 학회명, 개최지.', 'Presenter (Year). Title. Conference Name, Location.'))) }] },
      { t: 'section', title: T('연구 경력', 'Research Experience'), blocks: rows(2, null).map(() => ({
        t: 'entry', title: ph(T('연구 과제명 · 역할', 'Project Title · Role')), right: ph(T('YYYY.MM ~ YYYY.MM', 'YYYY – YYYY')),
        sub: ph(T('연구실 · 기관 · 지원 기관', 'Lab · Institution · Funding Agency')), bullets: [ph(T('맡은 연구 내용과 사용한 방법', 'What you studied and the methods you used'))] })) },
      { t: 'section', title: T('강의·조교 경력', 'Teaching Experience'), blocks: [{ t: 'entry', title: ph(T('과목명 · 조교/강사', 'Course · TA / Instructor')), right: ph('YYYY'), sub: ph(T('학교명', 'University')) }] },
      { t: 'section', title: T('연구비·수상', 'Grants & Awards'), blocks: [{ t: 'bullets', items: rows(2, ph(T('수상·장학명, 수여 기관 (연도)', 'Award or grant, Institution (Year)'))) }] },
      { t: 'section', title: T('연구 방법·도구', 'Skills'), blocks: [{ t: 'lines', items: [
        [T('연구 방법', 'Methods'), ph(T('예: 실험 설계, 회귀 분석', 'e.g., Experimental design, Regression'))],
        [T('소프트웨어', 'Software'), ph(T('예: R, Python, MATLAB', 'e.g., R, Python, MATLAB'))],
        [T('언어', 'Languages'), ph(T('예: 영어 (논문 작성 가능)', 'e.g., Korean (Native), English (Fluent)'))]] }] },
      { t: 'section', title: T('추천인', 'References'), blocks: [{ t: 'lines', items: rows(2, [ph(T('성함 · 직위', 'Name · Title')), ph(T('소속 · 이메일', 'Institution · Email'))]) }] },
    ];
  }

  function researchStatement(s, lang) {
    const T = (ko, en) => (lang === 'ko' ? ko : en);
    const inst = s.company || T('지원 기관', 'the Institution');
    return [
      { t: 'title', text: T('연구계획서', 'Research Statement'), sub: [ph(T('이름', 'Your Name')), ' · ', s.company || ph(T('지원 기관', 'Institution'))] },
      { t: 'section', title: T('연구 개요', 'Overview'), blocks: [{ t: 'para', text: ph(T('나의 핵심 연구 질문과 그것이 왜 중요한지 3~4줄로', '3–4 lines: your central research question and why it matters')) }] },
      { t: 'section', title: T('지금까지의 연구', 'Past Research'), blocks: [
        { t: 'para', text: ph(T('대표 연구 1: 문제 → 방법 → 결과 → 의미', 'Project 1: problem → method → findings → significance')) },
        { t: 'para', text: ph(T('대표 연구 2: 문제 → 방법 → 결과 → 의미', 'Project 2: problem → method → findings → significance')) }] },
      { t: 'section', title: T('앞으로의 연구 계획', 'Future Research'), blocks: [{ t: 'para', text: ph(T('단기(1~2년)·장기 계획과 기대 성과', 'Short-term (1–2 years) and long-term plans, with expected outcomes')) }] },
      { t: 'section', title: T(`${inst}과의 적합성`, `Fit with ${inst}`), blocks: [{ t: 'para', text: ph(T('함께 연구하고 싶은 교수·연구실, 활용할 장비·자원, 협업 가능성', 'Faculty or labs you want to work with, resources you will use, and possible collaborations')) }] },
      { t: 'section', title: T('참고문헌', 'References'), blocks: [{ t: 'bullets', numbered: true, items: rows(2, ph(T('저자 (연도). 제목. 학술지명.', 'Author (Year). Title. Journal.'))) }] },
    ];
  }

  function freeIntro(s, lang) {
    const T = (ko, en) => (lang === 'ko' ? ko : en);
    const portfolio = s.org === 'creative';
    const co = s.company || T('이 회사', 'This Company');
    const blocks = [
      { t: 'title', text: docName('free-intro', lang, s), sub: appliedLine(s, lang) },
      { t: 'section', title: T('한 줄 소개', 'In One Line'), blocks: [{ t: 'para', text: ph(T('나를 한 문장으로 (예: 데이터로 사용자 문제를 찾는 주니어 PM)', 'One sentence about you (e.g., A junior PM who finds user problems in data)')) }] },
      { t: 'section', title: T(`왜 ${co}인가요`, `Why ${co}`), blocks: [{ t: 'para', text: ph(T('회사의 제품·서비스·문화 중 끌린 점을 내 경험과 연결해서', 'Connect what draws you to the product or culture with your own experience')) }] },
      { t: 'section', title: T('이런 일을 해 봤어요', 'What I Have Done'), blocks: [{ t: 'bullets', items: rows(3, ph(T('경험 + 내가 한 일 + 결과(숫자)', 'Experience + your action + result in numbers'))) }] },
      { t: 'section', title: T('함께 일하면 기대할 수 있는 것', 'What I Bring'), blocks: [{ t: 'para', text: ph(T('입사 후 첫 6개월 안에 기여할 수 있는 일', 'What you can contribute in your first six months')) }] },
    ];
    if (portfolio) blocks.push({ t: 'section', title: T('포트폴리오', 'Portfolio'), blocks: [{ t: 'lines', items: [
      [T('링크', 'Link'), ph('URL')], [T('대표 작업', 'Highlights'), ph(T('작업명 – 내 역할 – 결과', 'Work – your role – result'))]] }] });
    return blocks;
  }

  const TEMPLATES = {
    'kr-table': (s) => krTable(s, false),
    'kr-blind': (s) => krTable(s, true),
    'ncs-exp': ncsExp,
    'en-resume': enResume,
    'career-desc': careerDesc,
    'creative-one': creativeOne,
    'academic-cv': academicCv,
    'kr-questions': (s) => krQuestions(s, 'corp'),
    'ncs': (s) => krQuestions(s, 'ncs'),
    'en-cover': enCover,
    'free-intro': freeIntro,
    'research-statement': researchStatement,
  };
  function buildDoc(id, lang, s) {
    return { id, lang, name: docName(id, lang, s), blocks: TEMPLATES[id](s, lang) };
  }

  // ── 미리보기(HTML) ──
  function parts(x) {
    return toParts(x).map((p) => p.ph ? h('span', { class: 'ph' }, `[${p.text}]`) : p.b ? h('strong', null, p.text) : document.createTextNode(p.text));
  }
  function renderBlocks(blocks, lang) { return blocks.map((b) => renderBlock(b, lang)); }
  function renderBlock(b, lang) {
    switch (b.t) {
      case 'title':
        return h('div', { class: 'd-title' }, h('h1', null, b.text), b.sub ? h('p', { class: 'd-title-sub' }, parts(b.sub)) : null);
      case 'header':
        return h('header', { class: 'd-header' + (b.align === 'center' ? ' center' : '') },
          h('div', { class: 'd-name' }, parts(b.name)),
          b.sub ? h('div', { class: 'd-role' }, parts(b.sub)) : null,
          b.contacts ? h('div', { class: 'd-contacts' }, b.contacts.map((c) => h('span', null, parts(c)))) : null);
      case 'section':
        return h('section', { class: 'd-sec' }, h('h2', null, b.title), renderBlocks(b.blocks, lang));
      case 'note':
        return h('p', { class: 'd-note' }, parts(b.text));
      case 'kv': {
        const four = b.rows.some((r) => r.length >= 4);
        return h('table', { class: 'd-kv' }, h('tbody', null, b.rows.map((r, ri) => h('tr', null,
          b.photo && ri === 0 ? h('td', { class: 'photo', rowspan: String(b.rows.length) }, lang === 'ko' ? '사진\n(3×4cm)' : 'Photo') : null,
          r.map((c, i) => i % 2 === 0
            ? h('th', { scope: 'row' }, parts(c))
            : h('td', { colspan: four && r.length === 2 ? '3' : null }, parts(c)))))));
      }
      case 'table':
        return h('table', { class: 'd-table' },
          h('colgroup', null, b.widths.map((w) => h('col', { style: { width: `${(w * 100).toFixed(1)}%` } }))),
          h('thead', null, h('tr', null, b.head.map((x) => h('th', { scope: 'col' }, x)))),
          h('tbody', null, b.rows.map((r) => h('tr', null, r.map((c) => h('td', null, parts(c)))))));
      case 'entry':
        return h('div', { class: 'd-entry' },
          h('div', { class: 'd-entry-top' }, h('span', { class: 'd-entry-title' }, parts(b.title)), b.right ? h('span', { class: 'd-entry-right' }, parts(b.right)) : null),
          b.sub ? h('div', { class: 'd-entry-sub' }, parts(b.sub)) : null,
          b.bullets && b.bullets.length ? h('ul', { class: 'd-bullets' }, b.bullets.map((x) => h('li', null, parts(x)))) : null);
      case 'bullets':
        return h(b.numbered ? 'ol' : 'ul', { class: 'd-bullets' }, b.items.map((x) => h('li', null, parts(x))));
      case 'lines':
        return h('div', { class: 'd-lines' + (b.stacked ? ' stacked' : '') }, b.items.map(([l, v]) => h('p', null, h('strong', null, parts(l)), b.stacked ? ' ' : ': ', parts(v))));
      case 'para':
        return h('p', { class: 'd-para' }, parts(b.text));
      case 'addr':
        return h('div', { class: 'd-addr' }, b.lines.map((l) => h('div', null, parts(l))));
      case 'question':
        return h('div', { class: 'd-q' },
          h('p', { class: 'd-q-title' }, h('span', { class: 'no' }, `${b.no}. `), parts(b.q), b.limit ? h('span', { class: 'limit' }, lang === 'ko' ? ` (${b.limit} 이내)` : ` (${b.limit})`) : null),
          h('div', { class: 'd-q-box' }, h('p', null, h('strong', null, parts(ph(b.sub)))), h('p', null, parts(ph(b.guide)))));
      case 'photo':
        return h('div', { class: 'd-photo' }, lang === 'ko' ? '사진' : 'Photo');
      case 'sign':
        return h('div', { class: 'd-sign' }, b.lines.map((l) => h('p', null, parts(l))));
      case 'twocol':
        return h('div', { class: 'd-two' }, h('aside', { class: 'd-side' }, renderBlocks(b.side, lang)), h('div', { class: 'd-main' }, renderBlocks(b.main, lang)));
      default:
        return null;
    }
  }
  function renderDoc(doc, s) {
    const el = h('article', { class: `doc ${s.tone} lang-${doc.lang}`, lang: doc.lang, 'aria-label': doc.name, 'data-doc': doc.id }, renderBlocks(doc.blocks, doc.lang));
    el.style.setProperty('--accent', ACCENTS[s.accent]);
    return el;
  }

  // ── 글(복사용) ──
  const textOf = (x) => toParts(x).map((p) => (p.ph ? `[${p.text}]` : p.text)).join('');
  function textify(blocks, lang, out = []) {
    for (const b of blocks) {
      switch (b.t) {
        case 'title': out.push(b.text); if (b.sub) out.push(textOf(b.sub)); out.push(''); break;
        case 'header':
          out.push(textOf(b.name)); if (b.sub) out.push(textOf(b.sub));
          if (b.contacts) out.push(b.contacts.map(textOf).join(' | '));
          out.push(''); break;
        case 'section': out.push(`■ ${b.title}`); textify(b.blocks, lang, out); out.push(''); break;
        case 'note': out.push(`※ ${textOf(b.text)}`); break;
        case 'kv':
          if (b.photo) out.push(lang === 'ko' ? '[사진 3×4cm]' : '[Photo]');
          b.rows.forEach((r) => {
            const pairs = [];
            for (let i = 0; i < r.length; i += 2) pairs.push(`${textOf(r[i])}: ${textOf(r[i + 1])}`);
            out.push(pairs.join('   '));
          });
          break;
        case 'table':
          out.push(b.head.join(' | '));
          b.rows.forEach((r) => out.push(r.map(textOf).join(' | ')));
          break;
        case 'entry':
          out.push(b.right ? `${textOf(b.title)}  —  ${textOf(b.right)}` : textOf(b.title));
          if (b.sub) out.push(textOf(b.sub));
          (b.bullets || []).forEach((x) => out.push(`  • ${textOf(x)}`));
          break;
        case 'bullets': b.items.forEach((x, i) => out.push(`  ${b.numbered ? `${i + 1}.` : '•'} ${textOf(x)}`)); break;
        case 'lines': b.items.forEach(([l, v]) => out.push(`${textOf(l)}: ${textOf(v)}`)); break;
        case 'para': out.push(textOf(b.text)); break;
        case 'addr': b.lines.forEach((l) => out.push(textOf(l))); out.push(''); break;
        case 'question':
          out.push(`${b.no}. ${textOf(b.q)}${b.limit ? (lang === 'ko' ? ` (${b.limit} 이내)` : ` (${b.limit})`) : ''}`);
          out.push(`   [${b.sub}]`, `   [${b.guide}]`, '');
          break;
        case 'photo': out.push(lang === 'ko' ? '[사진]' : '[Photo]'); break;
        case 'sign': b.lines.forEach((l) => out.push(textOf(l))); break;
        case 'twocol': textify(b.side, lang, out); textify(b.main, lang, out); break;
      }
    }
    return out;
  }

  function resultText(s, recs, rec) {
    const info = [
      `지원처 유형: ${ORG_LABEL[s.org]}`,
      s.company && `회사·기관: ${s.company}`,
      s.industry && `산업: ${s.industry}`,
      s.role && `직무: ${s.role}`,
      `언어: ${LANG_LABEL[s.lang]}`,
      `경력: ${s.career === 'exp' ? `경력${s.years ? ` ${s.years}년` : ''}` : '신입'}`,
    ].filter(Boolean).join(' · ');
    const out = ['[이력서·자기소개서 양식 추천 결과]', info, ''];
    recs.forEach((r, i) => {
      out.push(`${i + 1}. (${r.badge}) ${r.resumeName} + ${r.coverName} [${r.lang === 'ko' ? '국문' : '영문'}]`);
      r.why.forEach((w) => out.push(`   - ${w}`));
    });
    out.push('', '준비 팁');
    tipsFor(s, rec).forEach((t) => out.push(`   - ${t}`));
    const line = '────────────────────';
    const resume = buildDoc(rec.resume, rec.lang, s);
    const cover = buildDoc(rec.cover, rec.lang, s);
    out.push('', line, `선택한 양식: ${rec.resumeName} + ${rec.coverName}`, '[괄호] 안을 내 내용으로 바꿔 쓰면 돼요.', line, '');
    out.push(`《${resume.name}》`, '');
    textify(resume.blocks, rec.lang, out);
    out.push(line, '', `《${cover.name}》`, '');
    textify(cover.blocks, rec.lang, out);
    return out.join('\n').replace(/\n{3,}/g, '\n\n').trim() + '\n';
  }

  // ── 화면 갱신 ──
  const state = { recs: [], sig: '', selected: 0, tab: 'resume' };

  function update() {
    const s = readState();
    $('#years-wrap').hidden = s.career !== 'exp';
    const recs = recommend(s);
    const sig = recs.map((r) => r.key).join(',');
    if (sig !== state.sig) { state.sig = sig; state.selected = 0; }
    state.recs = recs;
    renderRecs(s);
  }

  function renderRecs(s) {
    const box = $('#recs');
    box.replaceChildren();
    $('#tips').replaceChildren();
    if (!state.recs.length) {
      box.append(h('div', { class: 'empty' }, '지원처 유형을 고르면 어울리는 양식을 추천해 드려요.'));
      $('#preview-panel').hidden = true;
      return;
    }
    const list = h('div', { class: 'rec-list' });
    state.recs.forEach((r, i) => {
      const sel = i === state.selected;
      list.append(h('article', { class: 'rec' + (sel ? ' is-selected' : ''), 'data-rec': String(i) },
        h('div', { class: 'rec-top' },
          h('span', { class: 'badge' + (i === 0 ? ' best' : '') }, r.badge),
          h('span', { class: 'chip' }, r.lang === 'ko' ? '국문' : '영문')),
        h('h3', null, `${r.resumeName} + ${r.coverName}`),
        h('p', { class: 'rec-docs' }, `이력서: ${r.resumeName} · 자기소개서: ${r.coverName}`),
        h('ul', null, r.why.map((w) => h('li', null, w))),
        h('div', { class: 'rec-btns' },
          h('button', { type: 'button', class: 'btn' + (sel ? ' ghost' : ''), 'aria-pressed': sel ? 'true' : 'false', 'data-action': 'preview', onclick: () => selectRec(i) }, sel ? '미리보는 중' : '미리보기'),
          h('button', { type: 'button', class: 'btn primary', 'data-action': 'decide', onclick: () => { selectRec(i); openFileDialog(); } }, '이 양식으로 결정'))));
    });
    box.append(list);
    const rec = state.recs[state.selected];
    $('#tips').append(h('div', { class: 'tips' }, h('h3', null, '준비 팁'), h('ul', null, tipsFor(s, rec).map((t) => h('li', null, t)))));
    renderPreview(s);
  }

  function selectRec(i) {
    state.selected = i;
    renderRecs(readState());
  }

  function renderPreview(s) {
    const rec = state.recs[state.selected];
    $('#preview-panel').hidden = false;
    const id = state.tab === 'resume' ? rec.resume : rec.cover;
    const doc = buildDoc(id, rec.lang, s);
    $('#tab-resume').textContent = `이력서 · ${rec.resumeName}`;
    $('#tab-cover').textContent = `자기소개서 · ${rec.coverName}`;
    $('#tab-resume').setAttribute('aria-selected', String(state.tab === 'resume'));
    $('#tab-cover').setAttribute('aria-selected', String(state.tab === 'cover'));
    $('#preview-caption').textContent = `${state.selected + 1}번 추천 양식 · [괄호] 안의 회색 글씨를 내 내용으로 바꿔 쓰면 돼요.`;
    $('#paper').replaceChildren(renderDoc(doc, s));
  }

  // ── 예시 채우기 ──
  const EXAMPLES = [
    { label: '외국계 마케팅 신입 (국문+영문)', org: 'foreign', company: 'Example Pharma Korea', industry: '제약·바이오', jobrole: 'Marketing Associate',
      doclang: 'both', career: 'new', years: '', tone: 'modern', accent: 'navy', photo: false, questions: '' },
    { label: '공공기관 사무행정 신입 (문항 2개 포함)', org: 'public', company: '한국○○공사', industry: '에너지 공공기관', jobrole: '사무행정',
      doclang: 'ko', career: 'new', years: '', tone: 'classic', accent: 'green', photo: true,
      questions: '지원 동기와 입사 후 포부를 기술해 주십시오. (600자)\n공동의 목표를 위해 다른 사람과 협력했던 경험을 기술해 주십시오. (800자)' },
    { label: 'IT 스타트업 백엔드 경력 3년', org: 'startup', company: '핀테크 스타트업', industry: '핀테크', jobrole: '백엔드 개발자',
      doclang: 'ko', career: 'exp', years: '3', tone: 'classic', accent: 'charcoal', photo: false, questions: '' },
  ];
  let exampleIdx = 0;
  function fillExample() {
    const ex = EXAMPLES[exampleIdx % EXAMPLES.length];
    exampleIdx++;
    for (const k of ['org', 'company', 'industry', 'jobrole', 'years', 'questions']) fv(k).value = ex[k];
    for (const k of ['doclang', 'career', 'tone', 'accent']) {
      form.querySelectorAll(`input[name="${k}"]`).forEach((r) => { r.checked = r.value === ex[k]; });
    }
    fv('photo').checked = ex.photo;
    state.tab = 'resume';
    $('#example-status').textContent = `예시 ${((exampleIdx - 1) % EXAMPLES.length) + 1}/${EXAMPLES.length}: ${ex.label}을(를) 채웠어요. 한 번 더 누르면 다른 예시가 나와요.`;
    update();
  }

  // ── 결과 복사 ──
  async function copyText(text) {
    try {
      if (navigator.clipboard && window.isSecureContext !== false) {
        await navigator.clipboard.writeText(text);
        return true;
      }
    } catch (e) { /* 아래 방법으로 다시 시도 */ }
    const ta = h('textarea', { readonly: true, 'aria-hidden': 'true' });
    ta.value = text;
    ta.style.setProperty('position', 'fixed');
    ta.style.setProperty('left', '-9999px');
    document.body.append(ta);
    ta.select();
    let ok = false;
    try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
    ta.remove();
    return ok;
  }
  let toastTimer;
  function toast(msg) {
    const t = $('#toast');
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove('show'), 2600);
  }

  // ── 파일 만들기 ──
  const dlg = $('#dlg-file');
  function openFileDialog() {
    const rec = state.recs[state.selected];
    if (!rec) return;
    $('#dlg-desc').textContent = `선택한 양식: ${rec.resumeName} + ${rec.coverName} (${rec.lang === 'ko' ? '국문' : '영문'})`;
    $('#dl-resume').textContent = `이력서 Word 파일 (.docx) · ${rec.resumeName}`;
    $('#dl-cover').textContent = `자기소개서 Word 파일 (.docx) · ${rec.coverName}`;
    $('#dlg-status').textContent = '';
    if (typeof dlg.showModal === 'function') dlg.showModal(); else dlg.setAttribute('open', '');
  }
  function closeDialog() { if (dlg.open) dlg.close(); }

  const safeName = (x) => x.replace(/[\\/:*?"<>|\u0000-\u001F]/g, '').replace(/\s+/g, '_').slice(0, 80);
  function downloadDocx(which) {
    const s = readState();
    const rec = state.recs[state.selected];
    const doc = buildDoc(which === 'resume' ? rec.resume : rec.cover, rec.lang, s);
    const blob = window.DocxGen.makeDocx(doc, { lang: rec.lang, tone: s.tone, accent: ACCENTS[s.accent] });
    const kind = which === 'resume' ? '이력서' : '자기소개서';
    const filename = safeName(`${kind}_${s.company || '양식'}_${doc.name}`) + '.docx';
    const url = URL.createObjectURL(blob);
    const a = h('a', { href: url, download: filename });
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
    $('#dlg-status').textContent = `${filename} 파일을 만들었어요. 다운로드 폴더를 확인하세요.`;
  }
  function printPdf() {
    const s = readState();
    const rec = state.recs[state.selected];
    const area = $('#print-area');
    area.replaceChildren(renderDoc(buildDoc(rec.resume, rec.lang, s), s), renderDoc(buildDoc(rec.cover, rec.lang, s), s));
    closeDialog();
    const clear = () => { area.replaceChildren(); window.removeEventListener('afterprint', clear); };
    window.addEventListener('afterprint', clear);
    window.print();
  }

  // ── 이벤트 연결 ──
  form.addEventListener('input', update);
  form.addEventListener('change', update);
  form.addEventListener('submit', (e) => e.preventDefault());
  form.addEventListener('reset', () => {
    $('#example-status').textContent = '';
    state.tab = 'resume';
    setTimeout(update, 0);
  });
  $('#btn-example').addEventListener('click', fillExample);
  document.querySelectorAll('.tabs [role="tab"]').forEach((b) => b.addEventListener('click', () => {
    state.tab = b.dataset.tab;
    renderPreview(readState());
  }));
  $('#btn-copy').addEventListener('click', async () => {
    const s = readState();
    const rec = state.recs[state.selected];
    if (!rec) return;
    const ok = await copyText(resultText(s, state.recs, rec));
    toast(ok ? '추천 결과와 선택한 양식을 복사했어요.' : '복사하지 못했어요. 브라우저의 클립보드 권한을 확인해 주세요.');
  });
  $('#btn-decide').addEventListener('click', openFileDialog);
  $('#dl-resume').addEventListener('click', () => downloadDocx('resume'));
  $('#dl-cover').addEventListener('click', () => downloadDocx('cover'));
  $('#dl-pdf').addEventListener('click', printPdf);
  $('#dlg-close').addEventListener('click', closeDialog);
  dlg.addEventListener('click', (e) => {
    if (e.target !== dlg) return;
    const r = dlg.getBoundingClientRect();
    const inside = e.clientX >= r.left && e.clientX <= r.right && e.clientY >= r.top && e.clientY <= r.bottom;
    if (!inside) closeDialog();
  });

  // 테스트·점검용으로 순수 함수만 노출 (네트워크와 무관)
  window.__resumeHelper = { recommend, parseQuestions, buildDoc, resultText, readState };

  update();
})();
