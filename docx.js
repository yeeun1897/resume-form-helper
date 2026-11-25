/* 브라우저 안에서 Word(.docx) 파일을 만드는 작은 생성기.
   외부 라이브러리나 서버 없이, 문서 구조(blocks)를 WordprocessingML로 바꾸고
   압축하지 않는(stored) ZIP으로 묶어요. */
(function (global) {
  'use strict';

  const enc = new TextEncoder();

  // ── ZIP (stored) ──
  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(u8) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < u8.length; i++) c = CRC_TABLE[(c ^ u8[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }
  function zip(files) {
    const now = new Date();
    const dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
    const dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
    const out = [], central = [];
    let offset = 0;
    for (const f of files) {
      const name = enc.encode(f.name);
      const data = typeof f.data === 'string' ? enc.encode(f.data) : f.data;
      const crc = crc32(data);
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true);
      lh.setUint16(4, 20, true);
      lh.setUint16(6, 0x0800, true);
      lh.setUint16(8, 0, true);
      lh.setUint16(10, dosTime, true);
      lh.setUint16(12, dosDate, true);
      lh.setUint32(14, crc, true);
      lh.setUint32(18, data.length, true);
      lh.setUint32(22, data.length, true);
      lh.setUint16(26, name.length, true);
      lh.setUint16(28, 0, true);
      out.push(new Uint8Array(lh.buffer), name, data);

      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true);
      ch.setUint16(4, 20, true);
      ch.setUint16(6, 20, true);
      ch.setUint16(8, 0x0800, true);
      ch.setUint16(10, 0, true);
      ch.setUint16(12, dosTime, true);
      ch.setUint16(14, dosDate, true);
      ch.setUint32(16, crc, true);
      ch.setUint32(20, data.length, true);
      ch.setUint32(24, data.length, true);
      ch.setUint16(28, name.length, true);
      ch.setUint32(42, offset, true);
      central.push(new Uint8Array(ch.buffer), name);
      offset += 30 + name.length + data.length;
    }
    const cdSize = central.reduce((a, b) => a + b.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true);
    end.setUint16(8, files.length, true);
    end.setUint16(10, files.length, true);
    end.setUint32(12, cdSize, true);
    end.setUint32(16, offset, true);
    return new Blob([...out, ...central, new Uint8Array(end.buffer)], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
  }

  // ── WordprocessingML 조각 ──
  const esc = (s) => String(s)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  const PAGE_W = 11906, PAGE_H = 16838, MARGIN_X = 1077, MARGIN_Y = 1020;
  const TEXT_W = PAGE_W - 2 * MARGIN_X;
  const GRAY = '5A6478', PH = '8A94A6', LINE = 'BFC5CE', FILL = 'F1F3F6';

  function run(text, o = {}) {
    let rpr = '';
    if (o.font) rpr += `<w:rFonts w:ascii="${esc(o.font)}" w:hAnsi="${esc(o.font)}" w:eastAsia="${esc(o.ea || o.font)}"/>`;
    if (o.b) rpr += '<w:b/>';
    if (o.i) rpr += '<w:i/>';
    if (o.color) rpr += `<w:color w:val="${o.color}"/>`;
    if (o.spacing) rpr += `<w:spacing w:val="${o.spacing}"/>`;
    if (o.sz) rpr += `<w:sz w:val="${o.sz}"/><w:szCs w:val="${o.sz}"/>`;
    const body = String(text).split('\n')
      .map((l, i) => (i ? '<w:br/>' : '') + `<w:t xml:space="preserve">${esc(l)}</w:t>`).join('');
    return `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}${body}</w:r>`;
  }
  const tabRun = () => '<w:r><w:tab/></w:r>';

  function para(runs, o = {}) {
    let ppr = '';
    if (o.keepNext) ppr += '<w:keepNext/>';
    if (o.border) ppr += `<w:pBdr>${o.border}</w:pBdr>`;
    if (o.tabRight) ppr += `<w:tabs><w:tab w:val="right" w:pos="${o.tabRight}"/></w:tabs>`;
    ppr += `<w:spacing w:before="${o.before ?? 0}" w:after="${o.after ?? 60}"${o.line ? ` w:line="${o.line}" w:lineRule="auto"` : ''}/>`;
    if (o.ind) ppr += `<w:ind w:left="${o.ind.left || 0}"${o.ind.hanging ? ` w:hanging="${o.ind.hanging}"` : ''}/>`;
    if (o.jc) ppr += `<w:jc w:val="${o.jc}"/>`;
    return `<w:p><w:pPr>${ppr}</w:pPr>${runs || ''}</w:p>`;
  }

  function cell(content, w, o = {}) {
    let pr = `<w:tcW w:w="${Math.round(w)}" w:type="dxa"/>`;
    if (o.span > 1) pr += `<w:gridSpan w:val="${o.span}"/>`;
    if (o.vMerge === 'restart') pr += '<w:vMerge w:val="restart"/>';
    else if (o.vMerge) pr += '<w:vMerge/>';
    if (o.fill) pr += `<w:shd w:val="clear" w:color="auto" w:fill="${o.fill}"/>`;
    if (o.mar) pr += `<w:tcMar><w:top w:w="${o.mar}" w:type="dxa"/><w:left w:w="${o.mar}" w:type="dxa"/><w:bottom w:w="${o.mar}" w:type="dxa"/><w:right w:w="${o.mar}" w:type="dxa"/></w:tcMar>`;
    pr += `<w:vAlign w:val="${o.vAlign || 'center'}"/>`;
    let c = content || '';
    if (!/<\/w:p>$/.test(c) && !/<w:p\/>$/.test(c)) c += '<w:p/>';
    return `<w:tc><w:tcPr>${pr}</w:tcPr>${c}</w:tc>`;
  }
  function row(cells, o = {}) {
    let pr = '';
    if (o.height) pr += `<w:trHeight w:val="${o.height}" w:hRule="atLeast"/>`;
    if (o.header) pr += '<w:tblHeader/>';
    return `<w:tr>${pr ? `<w:trPr>${pr}</w:trPr>` : ''}${cells.join('')}</w:tr>`;
  }
  function table(rows, widths, o = {}) {
    const total = Math.round(widths.reduce((a, b) => a + b, 0));
    const sides = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];
    const bd = o.noBorder
      ? sides.map((s) => `<w:${s} w:val="nil"/>`).join('')
      : sides.map((s) => `<w:${s} w:val="${o.dashed ? 'dashed' : 'single'}" w:sz="4" w:space="0" w:color="${o.borderColor || LINE}"/>`).join('');
    // tblInd: 표 테두리를 본문 여백·제목 밑줄과 맞춰요(셀 안쪽 여백만큼 안으로).
    return `<w:tbl><w:tblPr><w:tblW w:w="${total}" w:type="dxa"/><w:tblInd w:w="${o.noBorder ? 0 : 100}" w:type="dxa"/><w:tblBorders>${bd}</w:tblBorders><w:tblLayout w:type="fixed"/>` +
      `<w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr>` +
      `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${Math.round(w)}"/>`).join('')}</w:tblGrid>${rows.join('')}</w:tbl>`;
  }

  // 글 조각: 문자열 / {ph} 자리표시 / {b} 굵게 / 배열
  function toParts(x) {
    if (x == null || x === false) return [];
    if (Array.isArray(x)) return x.flatMap(toParts);
    if (typeof x === 'object') {
      if ('ph' in x) return [{ text: x.ph, ph: true }];
      if ('b' in x) return [{ text: x.b, b: true }];
    }
    return [{ text: String(x) }];
  }
  function runs(x, base = {}) {
    return toParts(x).map((p) => p.ph
      ? run(`[${p.text}]`, { ...base, color: PH })
      : run(p.text, p.b ? { ...base, b: true } : base)).join('');
  }

  // ── blocks → XML ──
  function build(blocks, ctx, W) {
    return blocks.map((b) => block(b, ctx, W)).join('');
  }

  function block(b, ctx, W) {
    const A = ctx.accent;
    const head = { font: ctx.headFont, ea: ctx.headEa };
    switch (b.t) {
      case 'title':
        return para(run(b.text, { ...head, b: true, sz: 36, spacing: 80, color: A }), { jc: 'center', after: b.sub ? 40 : 240 }) +
          (b.sub ? para(runs(b.sub, { color: GRAY, sz: 20 }), { jc: 'center', after: 240 }) : '');
      case 'header': {
        const jc = b.align === 'center' ? 'center' : undefined;
        let x = para(runs(b.name, { ...head, b: true, sz: 40, color: A }), { jc, after: 20 });
        if (b.sub) x += para(runs(b.sub, { sz: 23, color: GRAY }), { jc, after: 60 });
        if (b.contacts && b.contacts.length) {
          const sep = run('   |   ', { color: 'B7BFCB', sz: 18 });
          x += para(b.contacts.map((c) => runs(c, { sz: 18, color: GRAY })).join(sep), {
            jc, after: 200, border: `<w:bottom w:val="single" w:sz="10" w:space="6" w:color="${A}"/>`
          });
        } else {
          x += para('', { after: 120 });
        }
        return x;
      }
      case 'section': {
        const side = ctx.inSide;
        const border = side ? undefined : ctx.tone === 'modern'
          ? `<w:left w:val="single" w:sz="24" w:space="6" w:color="${A}"/>`
          : `<w:bottom w:val="single" w:sz="6" w:space="2" w:color="${A}"/>`;
        const upper = ctx.lang === 'en' && ctx.tone === 'classic' && !side;
        const title = para(run(upper ? b.title.toUpperCase() : b.title, {
          ...head, b: true, sz: side ? 19 : 22, color: A, spacing: upper || side ? 20 : 0
        }), { keepNext: true, before: side ? 200 : 260, after: 100, border, ind: ctx.tone === 'modern' && !side ? { left: 120 } : undefined });
        return title + build(b.blocks, ctx, W);
      }
      case 'note':
        return para(runs(b.text, { i: true, color: GRAY, sz: 17 }), { after: 100 });
      case 'kv': {
        const four = b.rows.some((r) => r.length >= 4);
        const photoW = b.photo ? 1700 : 0;
        const labW = 1450;
        const rest = W - photoW;
        const widths = four ? [labW, (rest - 2 * labW) / 2, labW, (rest - 2 * labW) / 2] : [labW, rest - labW];
        const grid = b.photo ? [photoW, ...widths] : widths;
        const rows = b.rows.map((r, ri) => {
          const cells = [];
          if (b.photo) {
            cells.push(ri === 0
              ? cell(para(run(ctx.lang === 'ko' ? '사진\n(3×4cm)' : 'Photo', { color: PH, sz: 18 }), { jc: 'center', after: 0 }), photoW, { vMerge: 'restart' })
              : cell('', photoW, { vMerge: true }));
          }
          for (let i = 0; i < r.length; i += 2) {
            const last = i + 2 >= r.length;
            const span = four && last && r.length === 2 ? 3 : 1;
            const vw = span === 3 ? widths.slice(1).reduce((a, c) => a + c, 0) : widths[i + 1];
            cells.push(cell(para(runs(r[i], { b: true, sz: 18 }), { after: 0 }), widths[i], { fill: FILL }));
            cells.push(cell(para(runs(r[i + 1], { sz: 18 }), { after: 0 }), vw, { span }));
          }
          return row(cells, { height: 440 });
        });
        return table(rows, grid) + para('', { after: 40 });
      }
      case 'table': {
        const widths = b.widths.map((f) => f * W);
        const headRow = row(b.head.map((hd, i) => cell(para(run(hd, { b: true, sz: 18 }), { jc: 'center', after: 0 }), widths[i], { fill: FILL })), { header: true, height: 380 });
        const body = b.rows.map((r) => row(r.map((c, i) => cell(para(runs(c, { sz: 18 }), { after: 0 }), widths[i])), { height: 440 }));
        return table([headRow, ...body], widths) + para('', { after: 40 });
      }
      case 'entry': {
        let x = para(runs(b.title, { b: true }) + (b.right ? tabRun() + runs(b.right, { color: GRAY, sz: 18 }) : ''), {
          tabRight: W, keepNext: true, before: 100, after: 20
        });
        if (b.sub) x += para(runs(b.sub, { i: true, color: GRAY, sz: 19 }), { after: 40, keepNext: !!(b.bullets && b.bullets.length) });
        if (b.bullets && b.bullets.length) x += block({ t: 'bullets', items: b.bullets }, ctx, W);
        return x;
      }
      case 'bullets':
        return b.items.map((it, i) => para(
          run(b.numbered ? `${i + 1}.` : '•', { color: A, b: !!b.numbered }) + tabRun() + runs(it),
          { ind: { left: 320, hanging: 320 }, after: 40 })).join('');
      case 'lines':
        return b.items.map(([label, value]) => b.stacked
          ? para(runs(label, { b: true, sz: 16, color: GRAY }), { after: 0 }) + para(runs(value, { sz: 18 }), { after: 100 })
          : para(runs(label, { b: true }) + run(':  ') + runs(value), { after: 50 })).join('');
      case 'para':
        return para(runs(b.text), { after: 140, line: 300 });
      case 'addr':
        return para(b.lines.map((l, i) => (i ? run('\n') : '') + runs(l)).join(''), { after: 200 });
      case 'question': {
        const lim = b.limit ? run(ctx.lang === 'ko' ? `  (${b.limit} 이내)` : `  (${b.limit})`, { color: GRAY, sz: 18 }) : '';
        const titleP = para(run(`${b.no}. `, { b: true, color: A }) + runs(b.q, { b: true }) + lim, { keepNext: true, before: 220, after: 80 });
        const inner = para(runs({ ph: b.sub }, { b: true }), { after: 80 }) + para(runs({ ph: b.guide }), { after: 0, line: 300 });
        return titleP + table([row([cell(inner, W, { vAlign: 'top', mar: 140 })], { height: 3200 })], [W]) + para('', { after: 40 });
      }
      case 'photo':
        return table([row([cell(para(run(ctx.lang === 'ko' ? '사진' : 'Photo', { color: PH, sz: 18 }), { jc: 'center', after: 0 }), 1500)], { height: 2000 })], [1500], { dashed: true }) + para('', { after: 120 });
      case 'sign':
        return para('', { after: 300 }) + b.lines.map((l, i) => para(runs(l), {
          jc: i === b.lines.length - 1 ? 'right' : 'center', after: 120, before: i === b.lines.length - 1 ? 200 : 0
        })).join('');
      case 'twocol': {
        const sideW = Math.round(W * 0.32), mainW = W - sideW;
        const sideCtx = { ...ctx, inSide: true };
        const sideXml = build(b.side, sideCtx, sideW - 360);
        const mainXml = build(b.main, ctx, mainW - 360);
        return table([row([
          cell(sideXml, sideW, { fill: 'F3F5F8', vAlign: 'top', mar: 180 }),
          cell(mainXml, mainW, { vAlign: 'top', mar: 180 })
        ], { height: 14000 })], [sideW, mainW], { noBorder: true });
      }
      default:
        return '';
    }
  }

  function makeDocx(doc, opts) {
    const lang = opts.lang;
    const ctx = {
      lang,
      tone: opts.tone,
      accent: opts.accent.replace('#', '').toUpperCase(),
      headFont: opts.tone === 'classic' ? (lang === 'en' ? 'Georgia' : '바탕') : (lang === 'en' ? 'Calibri' : '맑은 고딕'),
      headEa: opts.tone === 'classic' ? '바탕' : '맑은 고딕',
    };
    const body = build(doc.blocks, ctx, TEXT_W);
    const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}<w:sectPr><w:pgSz w:w="${PAGE_W}" w:h="${PAGE_H}"/><w:pgMar w:top="${MARGIN_Y}" w:right="${MARGIN_X}" w:bottom="${MARGIN_Y}" w:left="${MARGIN_X}" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr></w:body></w:document>`;
    const bodyFont = lang === 'en' ? 'Calibri' : '맑은 고딕';
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${bodyFont}" w:hAnsi="${bodyFont}" w:eastAsia="맑은 고딕" w:cs="${bodyFont}"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:val="${lang === 'en' ? 'en-US' : 'ko-KR'}" w:eastAsia="ko-KR"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="60" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>`;
    const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`;
    const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`;
    const docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
    const core = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${esc(doc.name)}</dc:title><dc:creator></dc:creator></cp:coreProperties>`;
    return zip([
      { name: '[Content_Types].xml', data: contentTypes },
      { name: '_rels/.rels', data: rels },
      { name: 'docProps/core.xml', data: core },
      { name: 'word/document.xml', data: documentXml },
      { name: 'word/_rels/document.xml.rels', data: docRels },
      { name: 'word/styles.xml', data: stylesXml },
    ]);
  }

  global.DocxGen = { makeDocx, toParts };
})(window);
