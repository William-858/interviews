/* 面試文件產生器：瀏覽器用 window.docx，Node 用 require('docx') */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('docx'));
  else root.DocGen = factory(root.docx);
})(typeof self !== 'undefined' ? self : this, function (D) {
  const {
    Document, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
    AlignmentType, BorderStyle, LevelFormat, Footer, PageNumber, HeadingLevel, TableLayoutType, LineRuleType,
  } = D;

  const FONT = 'Microsoft JhengHei';
  const ACCENT = '1F5C7A', LIGHT = 'E8F1F5', MUTED = '5A6B75', INK = '1A2A33';
  const ORG = '財團法人台灣水資源與農業研究院';
  const A4 = { width: 11906, height: 16838 };
  const L = v => Array.isArray(v) ? v : (v ? [v] : []);

  const run = (t, o = {}) => new TextRun({ text: String(t ?? ''), font: FONT, size: o.size || 19, bold: o.bold, color: o.color, italics: o.italics });
  function para(kids, o = {}) {
    if (!Array.isArray(kids)) kids = [run(kids, o)];
    return new Paragraph({
      children: kids, alignment: o.align, keepNext: o.keepNext, indent: o.indent, border: o.border, shading: o.shading,
      spacing: { before: o.before ?? 0, after: o.after ?? 50, line: o.line ?? 276, lineRule: LineRuleType.EXACT },
      numbering: o.bullet ? { reference: 'bullets', level: 0 } : undefined,
    });
  }
  const bullet = (t, o = {}) => para(Array.isArray(t) ? t : [run(t, o)], { ...o, bullet: true, after: o.after ?? 20 });
  function heading(t, o = {}) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_1, keepNext: true,
      children: [run(t, { size: o.size || 22, bold: true, color: ACCENT })],
      spacing: { before: o.before ?? 140, after: 60, line: 320, lineRule: LineRuleType.EXACT },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 2 } },
    });
  }
  const cb = { style: BorderStyle.SINGLE, size: 4, color: 'C5D3DA' };
  const borders = { top: cb, bottom: cb, left: cb, right: cb };
  function cell(c, w, o = {}) {
    const ps = L(c).map(x => x instanceof Paragraph ? x : para([run(x, { size: o.size || 18, bold: o.bold, color: o.color })], { after: 10, line: 264 }));
    return new TableCell({
      children: ps.length ? ps : [para('')], width: { size: w, type: WidthType.DXA }, borders,
      shading: o.fill ? { type: ShadingType.CLEAR, color: 'auto', fill: o.fill } : undefined,
      margins: { top: 40, bottom: 40, left: 80, right: 80 },
    });
  }
  function table(widths, rows, o = {}) {
    const total = widths.reduce((a, b) => a + b, 0);
    return new Table({
      width: { size: total, type: WidthType.DXA }, columnWidths: widths, layout: TableLayoutType.FIXED,
      rows: rows.map((r, i) => new TableRow({
        tableHeader: o.header && i === 0, cantSplit: true,
        children: r.map((c, j) => cell(c, widths[j], {
          fill: o.header && i === 0 ? ACCENT : (o.firstColFill && j === 0 ? LIGHT : undefined),
          color: o.header && i === 0 ? 'FFFFFF' : undefined,
          bold: (o.header && i === 0) || (o.firstColBold && j === 0), size: o.size,
        })),
      })),
    });
  }
  const numbering = { config: [{ reference: 'bullets', levels: [
    { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 340, hanging: 220 } } } },
  ] }] };
  const footer = t => new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
    run(t + '　｜　第 ', { size: 15, color: MUTED }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 15, color: MUTED }), run(' 頁', { size: 15, color: MUTED }),
  ] })] });

  /** 個人資訊列：姓名、角色、應徵職稱、面試日期（空白欄不顯示） */
  function personBlock(W, pairs) {
    const shown = pairs.filter(p => p[1]);
    const cells = [];
    shown.forEach(([k, v]) => { cells.push([k, v]); });
    const labelW = 1250;
    const n = cells.length;
    if (!n) return [];
    // 兩欄一列排列
    const rows = [];
    for (let i = 0; i < n; i += 2) {
      const a = cells[i], b = cells[i + 1];
      rows.push(b ? [a[0], a[1], b[0], b[1]] : [a[0], a[1], '', '']);
    }
    const valW = Math.round((W - labelW * 2) / 2);
    const widths = [labelW, valW, labelW, W - labelW * 2 - valW];
    return [new Table({
      width: { size: W, type: WidthType.DXA }, columnWidths: widths, layout: TableLayoutType.FIXED,
      rows: rows.map(r => new TableRow({ children: r.map((c, j) => cell(c, widths[j], {
        fill: j % 2 === 0 ? LIGHT : undefined, bold: j % 2 === 0 || j === 1, size: 19,
        color: j === 1 ? INK : undefined,
      })) })),
    }), para('', { after: 60 })];
  }
  function titleBlock(d, label) {
    return [
      para([run(label, { size: 17, bold: true, color: 'FFFFFF' })], { shading: { type: ShadingType.CLEAR, color: 'auto', fill: ACCENT }, after: 60 }),
      para([run(d.title, { size: 28, bold: true, color: INK })], { after: 30, line: 330 }),
      para([run(`委託機關：${d.client}　｜　計畫年度：${d.year}　｜　執行單位：${ORG}`, { size: 16, color: MUTED })], { after: 90 }),
    ];
  }

  /* ---------- 面試者版（≤2 頁） ---------- */
  function candidateDoc(p, person) {
    const c = p.candidate, W = A4.width - 2000;
    const body = [
      ...titleBlock(p, '  應徵者工作說明  '),
      ...personBlock(W, [['應徵者', person.name], ['應徵職稱', person.position], ['面試日期', person.date]]),
      para([run('職務定位　', { bold: true, color: ACCENT, size: 20 }), run(c.position_hint, { size: 20 })], {
        shading: { type: ShadingType.CLEAR, color: 'auto', fill: LIGHT }, after: 80, line: 290,
        border: { left: { style: BorderStyle.SINGLE, size: 18, color: ACCENT, space: 6 } },
      }),
      heading('一、這個計畫在做什麼', { before: 80 }),
      para(c.project_intro, { size: 20, after: 30, line: 290 }),
      para([run('工作的意義：', { bold: true, size: 20 }), run(c.why_it_matters, { size: 20 })], { after: 40, line: 290 }),
      heading('二、你加入後會做的工作'),
    ];
    c.main_tasks.forEach((t, i) => {
      body.push(para([run(`${i + 1}. ${t.name}　`, { bold: true, color: INK, size: 20 }), run(t.description, { size: 20 })], { after: 6, keepNext: true, line: 290 }));
      body.push(para([run('例：' + L(t.examples).join('；'), { size: 17, color: MUTED })], { after: 40, indent: { left: 280 } }));
    });
    body.push(heading('三、一年的工作節奏'));
    body.push(table([1700, W - 1700], c.annual_rhythm.map(r => [r.period, r.work]), { firstColFill: true, firstColBold: true }));
    const h = W / 2;
    body.push(heading('四、合作對象與參與產出的成果'));
    body.push(table([h, h], [['你會接觸的對象', '你會參與產出的成果'], [L(c.work_with).map(x => bullet(x, { size: 18 })), L(c.deliverables).map(x => bullet(x, { size: 18 }))]], { header: true }));
    body.push(heading('五、能力需求'));
    body.push(table([h, h], [['必備能力', '加分條件'], [L(c.skills_required).map(x => bullet(x, { size: 18 })), L(c.skills_plus).map(x => bullet(x, { size: 18 }))]], { header: true }));
    body.push(heading('六、工作型態'));
    L(c.work_style).forEach(x => body.push(bullet(x, { size: 19 })));
    body.push(heading('七、面試時請思考'));
    L(c.think_about).forEach((x, i) => body.push(para([run(`Q${i + 1}　`, { bold: true, color: ACCENT }), run(x)], { after: 30 })));
    body.push(para([run('本說明依本院過往同類計畫整理，實際工作內容以當年度契約及主管安排為準。', { size: 15, color: MUTED, italics: true })], { before: 100 }));
    return new Document({
      creator: ORG, title: `${p.title}｜應徵者工作說明｜${person.name}`, numbering,
      styles: { default: { document: { run: { font: FONT, size: 19 } } } },
      sections: [{ properties: { page: { size: A4, margin: { top: 800, bottom: 760, left: 1000, right: 1000 } } },
        footers: { default: footer(`${ORG}　應徵者工作說明`) }, children: body }],
    });
  }

  /* ---------- 面試官版（濃縮，≤4 頁） ---------- */
  function interviewerDoc(p, person, panel) {
    const s = p.interviewer_short, W = A4.width - 2 * 950;
    const body = [
      ...titleBlock(p, '  面試官準備資料｜內部使用，請勿交給應徵者  '),
      ...personBlock(W, [
        ['面試官', person.name + (person.lead ? '（主面試官）' : '')],
        ['面試小組', panel && panel.length > 1 ? panel.join('、') : ''],
        ['應徵職稱', person.position], ['面試日期', person.date],
      ]),
      heading('1. 計畫概要', { before: 40 }),
      para(s.overview, { after: 40 }),
      heading('2. 委託機關與利害關係人'),
      table([3000, W - 3000], [['對象', '角色／關係'], ...s.stakeholders.map(x => [x.who, x.role])], { header: true, firstColBold: true }),
      heading('3. 主要工作項目'),
      table([2300, W - 2300], s.work_items.map(x => [x.item, x.summary]), { firstColFill: true, firstColBold: true }),
      heading('4. 未來預計工作（推估，以新年度契約為準）'),
      table([Math.round(W * 0.58), W - Math.round(W * 0.58)], [['工作', '推估依據'], ...s.future.map(x => [x.item, x.basis])], { header: true }),
      heading('5. 新人日常會做的事'),
    ];
    const dw = L(s.daily_work), half = Math.ceil(dw.length / 2), hw = W / 2;
    const nb = { top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' } };
    body.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [hw, hw], layout: TableLayoutType.FIXED,
      borders: { top: nb.top, bottom: nb.top, left: nb.top, right: nb.top, insideHorizontal: nb.top, insideVertical: nb.top },
      rows: [new TableRow({ children: [dw.slice(0, half), dw.slice(half)].map(col => new TableCell({
        width: { size: hw, type: WidthType.DXA }, borders: { top: nb.top, bottom: nb.top, left: nb.top, right: nb.top },
        children: col.map(x => bullet(x, { size: 18 })) })) })] }));
    body.push(heading('6. 面試前要先懂的名詞'));
    body.push(table([2100, W - 2100], [['名詞', '說明'], ...s.key_terms.map(x => [x.term, x.explain])], { header: true, firstColBold: true }));
    body.push(heading('7. 要評估的能力與問法'));
    body.push(table([2300, W - 2300], [['能力', '怎麼問出來'], ...s.competencies.map(x => [x.competency, x.probe])], { header: true, firstColBold: true }));
    body.push(heading('8. 聽答案的訊號'));
    body.push(table([hw, hw], [['加分訊號（有料）', '警訊（需要追問）'], [L(s.green_flags).map(x => bullet(x, { size: 18 })), L(s.red_flags).map(x => bullet(x, { size: 18 }))]], { header: true }));
    body.push(heading('9. 建議題目'));
    s.questions.forEach((q, i) => {
      body.push(para([run(`Q${i + 1}. ${q.question}`, { bold: true, size: 19, color: INK }), run(`　${q.type || ''}｜${q.level || ''}`, { size: 15, color: MUTED })], { before: 70, after: 10, keepNext: true }));
      L(q.answer_points).forEach(a => body.push(bullet(a, { size: 18, after: 6 })));
      if (q.rubric) body.push(para([run('評分觀察：', { bold: true, size: 17, color: ACCENT }), run(q.rubric, { size: 17 })], { after: 20, indent: { left: 120 } }));
    });
    if (L(s.caveats).length) {
      body.push(heading('10. 資料注意事項'));
      L(s.caveats).forEach(x => body.push(bullet(x, { size: 17 })));
    }
    return new Document({
      creator: ORG, title: `${p.title}｜面試官準備資料｜${person.name}`, numbering,
      styles: { default: { document: { run: { font: FONT, size: 19 } } } },
      sections: [{ properties: { page: { size: A4, margin: { top: 760, bottom: 720, left: 950, right: 950 } } },
        footers: { default: footer(`${ORG}　面試官準備資料｜${person.name}｜內部使用`) }, children: body }],
    });
  }

  const safe = s => String(s || '').replace(/[\\/:*?"<>|]/g, '_').trim();
  const fileName = (kind, person, p) => `${kind}_${safe(person.name)}_${p.code}${p.slug}.docx`;

  return { candidateDoc, interviewerDoc, fileName, Packer: D.Packer };
});
