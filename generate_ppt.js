const pptxgen = require('pptxgenjs');
const path = require('path');

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5 inches
pres.author = 'Antigravity AI';
pres.company = 'TSMC Earnings Analysis';
pres.subject = '台積電 2026 Q3 法說會實績與 Q2/Q3 指標達成率比對報告';
pres.title = '台積電 (2330) 2026 Q3 法說會實績與指標達成率分析';

// Color Palette
const COLORS = {
  navy: '0F172A',
  darkBlue: '1E293B',
  primary: '0284C7',
  accent: '0EA5E9',
  gold: 'F59E0B',
  green: '10B981',
  greenBg: 'ECFDF5',
  blueBg: 'F0F9FF',
  amberBg: 'FFFBEB',
  lightBg: 'F8FAFC',
  white: 'FFFFFF',
  textMain: '1E293B',
  textMuted: '64748B',
  border: 'E2E8F0'
};

const FONT = 'Microsoft JhengHei';

function addHeader(slide, titleText, subtitleText, pageNum) {
  slide.background = { color: COLORS.lightBg };
  // Top bar
  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 13.33, h: 1.05,
    fill: { color: COLORS.navy }
  });
  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 1.05, w: 13.33, h: 0.06,
    fill: { color: COLORS.accent }
  });
  slide.addText(titleText, {
    x: 0.6, y: 0.15, w: 10.5, h: 0.45,
    fontFace: FONT, fontSize: 22, bold: true, color: COLORS.white
  });
  slide.addText(subtitleText, {
    x: 0.6, y: 0.6, w: 10.5, h: 0.35,
    fontFace: FONT, fontSize: 12, color: '94A3B8'
  });
  // Footer
  slide.addText('資料來源：台積電 2026 Q1–Q3 法人說明會 & NotebookLM 筆記本資料庫交叉比對', {
    x: 0.6, y: 7.1, w: 9.0, h: 0.3,
    fontFace: FONT, fontSize: 9.5, color: COLORS.textMuted
  });
  slide.addText(`Page ${pageNum} / 6`, {
    x: 11.2, y: 7.1, w: 1.5, h: 0.3,
    fontFace: FONT, fontSize: 9.5, align: 'right', color: COLORS.textMuted, bold: true
  });
}

// ==========================================================
// SLIDE 1: Cover Page
// ==========================================================
const slide1 = pres.addSlide();
slide1.background = { color: COLORS.navy };

// Decorative accent shapes
slide1.addShape(pres.ShapeType.rect, {
  x: 0, y: 0, w: 0.35, h: 7.5,
  fill: { color: COLORS.accent }
});
slide1.addShape(pres.ShapeType.rect, {
  x: 0.35, y: 0, w: 0.1, h: 7.5,
  fill: { color: COLORS.gold }
});

// Badge
slide1.addShape(pres.ShapeType.roundRect, {
  x: 1.0, y: 1.1, w: 4.2, h: 0.45,
  fill: { color: '1E3A8A' },
  line: { color: COLORS.accent, width: 1 }
});
slide1.addText('TSMC 2330.TW / TSM  |  2026 Q3 法說會深度追蹤', {
  x: 1.0, y: 1.1, w: 4.2, h: 0.45,
  fontFace: FONT, fontSize: 12, bold: true, color: '38BDF8', align: 'center'
});

// Main Title
slide1.addText('台積電 2026 Q3 法說會實績\n與 Q2/Q3 指標達成率比對報告', {
  x: 1.0, y: 1.8, w: 11.2, h: 1.8,
  fontFace: FONT, fontSize: 36, bold: true, color: COLORS.white, lineSpacingMultiple: 1.2
});

// Subtitle
slide1.addText('結合 NotebookLM 筆記本內部財測紀錄（Q1/Q2 法說會、7月營收、凱基報告）與 2026/10/15 最新法說會數據交叉驗證', {
  x: 1.0, y: 3.7, w: 11.2, h: 0.6,
  fontFace: FONT, fontSize: 15, color: 'CBD5E1'
});

// 4 KPI Summary Cards on Cover
const coverKpis = [
  { label: '3Q26 美元營收 (達成率 103.4%)', value: 'US$467.2 億', sub: '飛越財測高標 458 億美元 (+16.2% QoQ)', color: COLORS.accent },
  { label: '3Q26 毛利率 / 營益率', value: '66.4% / 58.7%', sub: '營益率飛越高標 58% (達成率 103.0%)', color: COLORS.green },
  { label: '3Q26 每股盈餘 (EPS)', value: 'NT$29.43 元', sub: '稅後純益 7,630.6 億元，雙創歷史新高', color: COLORS.gold },
  { label: '2026 全年美元營收成長率', value: '三度上修至 ~44%', sub: '由年初 ~30% → >30% → >40% → ~44%', color: 'A855F7' }
];

coverKpis.forEach((kpi, i) => {
  const xPos = 1.0 + i * 2.9;
  slide1.addShape(pres.ShapeType.roundRect, {
    x: xPos, y: 4.7, w: 2.7, h: 1.85,
    fill: { color: COLORS.darkBlue },
    line: { color: kpi.color, width: 1.5 }
  });
  slide1.addText(kpi.label, {
    x: xPos + 0.15, y: 4.85, w: 2.4, h: 0.35,
    fontFace: FONT, fontSize: 10.5, bold: true, color: '94A3B8'
  });
  slide1.addText(kpi.value, {
    x: xPos + 0.15, y: 5.25, w: 2.4, h: 0.55,
    fontFace: FONT, fontSize: 21, bold: true, color: COLORS.white
  });
  slide1.addText(kpi.sub, {
    x: xPos + 0.15, y: 5.85, w: 2.4, h: 0.55,
    fontFace: FONT, fontSize: 9.5, color: 'CBD5E1'
  });
});

slide1.addText('報告產出日期：2026 年 10 月 18 日   |   資料基準：台積電 3Q26 法說會 (2026/10/15) & NotebookLM 知識庫', {
  x: 1.0, y: 6.85, w: 11.2, h: 0.3,
  fontFace: FONT, fontSize: 10, color: '64748B'
});

// ==========================================================
// SLIDE 2: 3Q26 Achievement Rate Analysis
// ==========================================================
const slide2 = pres.addSlide();
addHeader(
  slide2,
  '一、 2026 第三季 (3Q26) 核心指標達成率總表',
  '對比 NotebookLM 筆記本內載之 3Q26 財測指引（2026年7月法說會）與 10/15 公布之實際財報',
  2
);

const table2Rows = [
  [
    { text: '財務指標項目', options: { bold: true, fill: COLORS.navy, color: COLORS.white, align: 'left' } },
    { text: 'NotebookLM 內載 3Q26 財測\n(7月法說會指引)', options: { bold: true, fill: COLORS.navy, color: COLORS.white, align: 'center' } },
    { text: '3Q26 法說會實際數\n(2026/10/15 公布)', options: { bold: true, fill: COLORS.navy, color: COLORS.white, align: 'center' } },
    { text: '對比中位數\n達成率', options: { bold: true, fill: COLORS.navy, color: COLORS.white, align: 'center' } },
    { text: '對比高標\n達成率', options: { bold: true, fill: COLORS.navy, color: COLORS.white, align: 'center' } },
    { text: '達成表現與關鍵評析', options: { bold: true, fill: COLORS.navy, color: COLORS.white, align: 'left' } }
  ],
  [
    { text: '美元合併營收', options: { bold: true, fill: COLORS.white } },
    { text: '446 ~ 458 億美元\n(中位數 452 億)', options: { fill: COLORS.white, align: 'center' } },
    { text: '467.2 億美元\n(季增16.2% / 年增41.2%)', options: { bold: true, fill: COLORS.greenBg, color: '065F46', align: 'center' } },
    { text: '103.36%', options: { bold: true, fill: COLORS.greenBg, color: '065F46', align: 'center' } },
    { text: '102.01%', options: { bold: true, fill: COLORS.greenBg, color: '065F46', align: 'center' } },
    { text: '飛越財測高標 +9.2 億美元，AI 與 N2/N3 需求超乎預期', options: { fill: COLORS.white } }
  ],
  [
    { text: '新台幣合併營收', options: { bold: true, fill: 'F1F5F9' } },
    { text: '1.427兆 ~ 1.465兆元\n(中位數 1.446 兆)', options: { fill: 'F1F5F9', align: 'center' } },
    { text: '1 兆 4,942.4 億元\n(季增17.6% / 年增50.9%)', options: { bold: true, fill: COLORS.greenBg, color: '065F46', align: 'center' } },
    { text: '103.31%', options: { bold: true, fill: COLORS.greenBg, color: '065F46', align: 'center' } },
    { text: '101.95%', options: { bold: true, fill: COLORS.greenBg, color: '065F46', align: 'center' } },
    { text: '飛越財測高標 +286 億元，單季營收逼近 1.5 兆大關', options: { fill: 'F1F5F9' } }
  ],
  [
    { text: '毛利率 (Gross Margin)', options: { bold: true, fill: COLORS.white } },
    { text: '65.0% ~ 67.0%\n(中位數 66.0%)', options: { fill: COLORS.white, align: 'center' } },
    { text: '66.4%\n(季減 1.3 ppts)', options: { bold: true, fill: COLORS.blueBg, color: '0369A1', align: 'center' } },
    { text: '100.61%\n(+0.4 ppt)', options: { bold: true, fill: COLORS.blueBg, color: '0369A1', align: 'center' } },
    { text: '99.10%\n(-0.6 ppt)', options: { fill: COLORS.white, align: 'center' } },
    { text: '超越財測中位數；抵銷 2nm (N2) 初期量產成本稀釋效應', options: { fill: COLORS.white } }
  ],
  [
    { text: '營業利益率 (OPM)', options: { bold: true, fill: 'F1F5F9' } },
    { text: '56.0% ~ 58.0%\n(中位數 57.0%)', options: { fill: 'F1F5F9', align: 'center' } },
    { text: '58.7%\n(季減 1.6 ppts)', options: { bold: true, fill: COLORS.greenBg, color: '065F46', align: 'center' } },
    { text: '102.98%\n(+1.7 ppts)', options: { bold: true, fill: COLORS.greenBg, color: '065F46', align: 'center' } },
    { text: '101.21%\n(+0.7 ppt)', options: { bold: true, fill: COLORS.greenBg, color: '065F46', align: 'center' } },
    { text: '飛越財測高標 58.0%，展現極佳營運槓桿與費用控管', options: { fill: 'F1F5F9' } }
  ],
  [
    { text: '美元兌新台幣匯率', options: { bold: true, fill: COLORS.white } },
    { text: '32.00 (假設基準)', options: { fill: COLORS.white, align: 'center' } },
    { text: '31.98 (實際平均)', options: { fill: COLORS.white, align: 'center' } },
    { text: '99.94%', options: { fill: COLORS.white, align: 'center' } },
    { text: '—', options: { fill: COLORS.white, align: 'center' } },
    { text: '實際匯率與財測假設幾乎完全一致，無匯兌失真影響', options: { fill: COLORS.white } }
  ],
  [
    { text: '稅後純益 / 每股盈餘', options: { bold: true, fill: 'F1F5F9' } },
    { text: '未提供官方區間\n(2Q26 實績 EPS 27.25元)', options: { fill: 'F1F5F9', align: 'center' } },
    { text: '7,630.6 億元\nEPS 29.43 元', options: { bold: true, fill: COLORS.amberBg, color: 'B45309', align: 'center' } },
    { text: '季增 +8.0%', options: { bold: true, fill: COLORS.amberBg, color: 'B45309', align: 'center' } },
    { text: '續創歷史新高', options: { bold: true, fill: COLORS.amberBg, color: 'B45309', align: 'center' } },
    { text: '單季獲利逼近 8,000 億元，前三季累計 EPS 達 78.76 元', options: { fill: 'F1F5F9' } }
  ]
];

slide2.addTable(table2Rows, {
  x: 0.6, y: 1.3, w: 12.13,
  colW: [1.8, 2.2, 2.2, 1.25, 1.25, 3.43],
  fontFace: FONT, fontSize: 10.5,
  border: { type: 'solid', color: COLORS.border, pt: 0.75 },
  valign: 'middle',
  margin: [5, 6, 5, 6]
});

// Bottom Insight Callout Box
slide2.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 5.8, w: 12.13, h: 1.05,
  fill: { color: COLORS.blueBg },
  line: { color: COLORS.primary, width: 1.2 }
});
slide2.addText(
  '📌 NotebookLM 筆記本「7月營收」進度交叉驗證：\n' +
  '• 筆記本來源《台積電七月營收優於預期》記錄：2026年7月營收為 NT$4,675.8 億元（MoM +5.6%, YoY +44.7%），當時已達成 Q3 財測中位數的 32.3%。\n' +
  '• 最終 3Q26 實績達 NT$1 兆 4,942.4 億元，換算 8、9 月平均單月營收高達 NT$5,133.3 億元，驗證下半年 AI 伺服器與蘋果新機備貨動能在季末加速衝高！',
  {
    x: 0.8, y: 5.85, w: 11.7, h: 0.95,
    fontFace: FONT, fontSize: 11, color: COLORS.textMain, lineSpacingMultiple: 1.25
  }
);

// ==========================================================
// SLIDE 3: 2Q26 Achievement Review & 1Q-2Q-3Q Trajectory
// ==========================================================
const slide3 = pres.addSlide();
addHeader(
  slide3,
  '二、 第二季 (2Q26) 財測達成率回顧與三季成長軌跡',
  '比對 NotebookLM 筆記本內《1Q26 Presentation》之 Q2 財測、Q2 實際表現與 Q3 最新季增幅度',
  3
);

const table3Rows = [
  [
    { text: '核心指標', options: { bold: true, fill: COLORS.navy, color: COLORS.white } },
    { text: '1Q26 實際數\n(筆記本 PDF 實績)', options: { bold: true, fill: COLORS.navy, color: COLORS.white, align: 'center' } },
    { text: '2Q26 財測指引\n(筆記本 4月法說紀錄)', options: { bold: true, fill: COLORS.navy, color: COLORS.white, align: 'center' } },
    { text: '2Q26 實際數\n(7月法說公布)', options: { bold: true, fill: COLORS.navy, color: COLORS.white, align: 'center' } },
    { text: '2Q26 達成率\n(中位數 / 高標)', options: { bold: true, fill: COLORS.navy, color: COLORS.white, align: 'center' } },
    { text: '3Q26 實際數\n(10/15 法說公布)', options: { bold: true, fill: COLORS.navy, color: COLORS.white, align: 'center' } },
    { text: '3Q26 vs 2Q26\n季增減 (QoQ)', options: { bold: true, fill: COLORS.navy, color: COLORS.white, align: 'center' } }
  ],
  [
    { text: '美元營收 (US$)', options: { bold: true, fill: COLORS.white } },
    { text: '359.0 億美元\n(財測 346-358億)', options: { fill: COLORS.white, align: 'center' } },
    { text: '390 ~ 402 億美元\n(中位數 396 億)', options: { fill: COLORS.white, align: 'center' } },
    { text: '402.0 億美元\n(達財測最高標)', options: { bold: true, fill: COLORS.blueBg, align: 'center' } },
    { text: '101.52% (中)\n100.00% (高)', options: { bold: true, fill: COLORS.greenBg, color: '065F46', align: 'center' } },
    { text: '467.2 億美元', options: { bold: true, fill: COLORS.amberBg, align: 'center' } },
    { text: '+16.2%\n(+65.2 億美元)', options: { bold: true, fill: COLORS.greenBg, color: '065F46', align: 'center' } }
  ],
  [
    { text: '台幣營收 (NT$)', options: { bold: true, fill: 'F1F5F9' } },
    { text: '1 兆 1,341.0 億元', options: { fill: 'F1F5F9', align: 'center' } },
    { text: '1.236兆 ~ 1.274兆元\n(以匯率 31.7 折算)', options: { fill: 'F1F5F9', align: 'center' } },
    { text: '1 兆 2,703.8 億元', options: { bold: true, fill: COLORS.blueBg, align: 'center' } },
    { text: '101.20% (中)\n99.69% (高)', options: { bold: true, fill: COLORS.greenBg, color: '065F46', align: 'center' } },
    { text: '1 兆 4,942.4 億元', options: { bold: true, fill: COLORS.amberBg, align: 'center' } },
    { text: '+17.6%\n(+2,238.6 億元)', options: { bold: true, fill: COLORS.greenBg, color: '065F46', align: 'center' } }
  ],
  [
    { text: '毛利率 (GM)', options: { bold: true, fill: COLORS.white } },
    { text: '66.2%\n(財測 63-65%)', options: { fill: COLORS.white, align: 'center' } },
    { text: '65.5% ~ 67.5%\n(中位數 66.5%)', options: { fill: COLORS.white, align: 'center' } },
    { text: '67.7%\n(超越財測高標)', options: { bold: true, fill: COLORS.blueBg, align: 'center' } },
    { text: '101.80% (中)\n100.30% (高)', options: { bold: true, fill: COLORS.greenBg, color: '065F46', align: 'center' } },
    { text: '66.4%', options: { bold: true, fill: COLORS.amberBg, align: 'center' } },
    { text: '-1.3 ppts\n(N2 初期量產稀釋)', options: { fill: COLORS.white, align: 'center' } }
  ],
  [
    { text: '營業利益率 (OPM)', options: { bold: true, fill: 'F1F5F9' } },
    { text: '58.1%\n(財測 54-56%)', options: { fill: 'F1F5F9', align: 'center' } },
    { text: '56.5% ~ 58.5%\n(中位數 57.5%)', options: { fill: 'F1F5F9', align: 'center' } },
    { text: '60.3%\n(大超財測高標)', options: { bold: true, fill: COLORS.blueBg, align: 'center' } },
    { text: '104.87% (中)\n103.08% (高)', options: { bold: true, fill: COLORS.greenBg, color: '065F46', align: 'center' } },
    { text: '58.7%', options: { bold: true, fill: COLORS.amberBg, align: 'center' } },
    { text: '-1.6 ppts\n(仍優於Q3財測高標)', options: { fill: 'F1F5F9', align: 'center' } }
  ],
  [
    { text: '稅後純益 / EPS', options: { bold: true, fill: COLORS.white } },
    { text: '5,724.8 億元\nEPS 22.08 元', options: { fill: COLORS.white, align: 'center' } },
    { text: '未提供官方區間', options: { fill: COLORS.white, align: 'center' } },
    { text: '7,065.6 億元\nEPS 27.25 元', options: { bold: true, fill: COLORS.blueBg, align: 'center' } },
    { text: 'QoQ +23.4%', options: { bold: true, fill: COLORS.greenBg, color: '065F46', align: 'center' } },
    { text: '7,630.6 億元\nEPS 29.43 元', options: { bold: true, fill: COLORS.amberBg, align: 'center' } },
    { text: '+8.0%\n(連三季創歷史新高)', options: { bold: true, fill: COLORS.greenBg, color: '065F46', align: 'center' } }
  ]
];

slide3.addTable(table3Rows, {
  x: 0.6, y: 1.3, w: 12.13,
  colW: [1.55, 1.75, 1.85, 1.75, 1.65, 1.75, 1.83],
  fontFace: FONT, fontSize: 10,
  border: { type: 'solid', color: COLORS.border, pt: 0.75 },
  valign: 'middle',
  margin: [6, 5, 6, 5]
});

// Two analytical cards at bottom of Slide 3
slide3.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 5.35, w: 5.9, h: 1.5,
  fill: { color: COLORS.white },
  line: { color: COLORS.accent, width: 1.2 }
});
slide3.addText('💡 2Q26 → 3Q26 營收與獲利動能解析', {
  x: 0.8, y: 5.45, w: 5.5, h: 0.35,
  fontFace: FONT, fontSize: 13, bold: true, color: COLORS.primary
});
slide3.addText(
  '• 連續三季超越財測上緣：1Q26 ($35.9B vs 35.8B)、2Q26 ($40.2B vs 40.2B)、3Q26 ($46.72B vs 45.8B) 皆達陣或飛越財測高標。\n' +
  '• 2026 前三季累計營收達 NT$3 兆 8,987.2 億元，累計稅後純益達 NT$2 兆 0,421.0 億元，前三季累計 EPS 高達 78.76 元！',
  {
    x: 0.8, y: 5.82, w: 5.5, h: 0.95,
    fontFace: FONT, fontSize: 10.5, color: COLORS.textMain, lineSpacingMultiple: 1.25
  }
);

slide3.addShape(pres.ShapeType.roundRect, {
  x: 6.83, y: 5.35, w: 5.9, h: 1.5,
  fill: { color: COLORS.white },
  line: { color: COLORS.gold, width: 1.2 }
});
slide3.addText('🔍 毛利率微調 (-1.3 ppts) 背後之結構原因', {
  x: 7.03, y: 5.45, w: 5.5, h: 0.35,
  fontFace: FONT, fontSize: 13, bold: true, color: 'B45309'
});
slide3.addText(
  '• 筆記本《鏡報》記載 Q1 法說會預告 Q2 毛利率衝高，實際 2Q26 毛利率果真創下 67.7% 驚人紀錄（超越高標 67.5%）。\n' +
  '• 3Q26 毛利率為 66.4%（仍高於財測中位數 66.0%），季減 1.3 ppts 完全符合管理層預期，主因 2 奈米 (N2) 進入初期量產爬坡帶來之折舊稀釋。',
  {
    x: 7.03, y: 5.82, w: 5.5, h: 0.95,
    fontFace: FONT, fontSize: 10.5, color: COLORS.textMain, lineSpacingMultiple: 1.25
  }
);

// ==========================================================
// SLIDE 4: Process Node & Platform Mix Evolution
// ==========================================================
const slide4 = pres.addSlide();
addHeader(
  slide4,
  '三、 先進製程與應用平台結構演變 (1Q26 → 2Q26 → 3Q26)',
  '2 奈米 (N2) 於第三季強勢放量佔比跳升至 8%，HPC 營收佔比突破三分之二達 68%',
  4
);

// Left Card: Process Node Table
slide4.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.3, w: 5.9, h: 5.55,
  fill: { color: COLORS.white },
  line: { color: COLORS.border, width: 1 }
});
slide4.addText('🔬 晶圓銷售製程佔比演變 (By Technology)', {
  x: 0.8, y: 1.45, w: 5.5, h: 0.4,
  fontFace: FONT, fontSize: 14, bold: true, color: COLORS.navy
});

const nodeRows = [
  [
    { text: '製程節點', options: { bold: true, fill: COLORS.navy, color: COLORS.white } },
    { text: '1Q26 佔比\n(筆記本 PDF)', options: { bold: true, fill: COLORS.navy, color: COLORS.white, align: 'center' } },
    { text: '2Q26 佔比\n(初期導入)', options: { bold: true, fill: COLORS.navy, color: COLORS.white, align: 'center' } },
    { text: '3Q26 佔比\n(最新法說)', options: { bold: true, fill: COLORS.navy, color: COLORS.white, align: 'center' } },
    { text: '趨勢變化', options: { bold: true, fill: COLORS.navy, color: COLORS.white, align: 'center' } }
  ],
  [
    { text: '2 奈米 (N2)', options: { bold: true, fill: COLORS.amberBg, color: 'B45309' } },
    { text: '0%', options: { fill: COLORS.white, align: 'center' } },
    { text: '~3%', options: { fill: COLORS.white, align: 'center' } },
    { text: '8%', options: { bold: true, fill: COLORS.amberBg, color: 'B45309', align: 'center' } },
    { text: '🚀 放量爬坡極快', options: { bold: true, fill: COLORS.amberBg, color: 'B45309', align: 'center' } }
  ],
  [
    { text: '3 奈米 (N3)', options: { bold: true, fill: COLORS.white } },
    { text: '25%', options: { fill: COLORS.white, align: 'center' } },
    { text: '~30%', options: { fill: COLORS.white, align: 'center' } },
    { text: '33%', options: { bold: true, fill: COLORS.blueBg, color: '0369A1', align: 'center' } },
    { text: '📈 躍居第一大主力', options: { fill: COLORS.white, align: 'center' } }
  ],
  [
    { text: '5 奈米 (N5)', options: { bold: true, fill: 'F1F5F9' } },
    { text: '36%', options: { fill: 'F1F5F9', align: 'center' } },
    { text: '~32%', options: { fill: 'F1F5F9', align: 'center' } },
    { text: '28%', options: { bold: true, fill: 'F1F5F9', align: 'center' } },
    { text: '維持滿載穩健貢獻', options: { fill: 'F1F5F9', align: 'center' } }
  ],
  [
    { text: '7 奈米 (N7)', options: { bold: true, fill: COLORS.white } },
    { text: '13%', options: { fill: COLORS.white, align: 'center' } },
    { text: '~11%', options: { fill: COLORS.white, align: 'center' } },
    { text: '10%', options: { fill: COLORS.white, align: 'center' } },
    { text: '比重隨先進製程稀釋', options: { fill: COLORS.white, align: 'center' } }
  ],
  [
    { text: '先進製程合計 (≤7nm)', options: { bold: true, fill: COLORS.greenBg, color: '065F46' } },
    { text: '74%', options: { bold: true, fill: COLORS.greenBg, align: 'center' } },
    { text: '~76%', options: { bold: true, fill: COLORS.greenBg, align: 'center' } },
    { text: '79%', options: { bold: true, fill: COLORS.greenBg, color: '065F46', align: 'center' } },
    { text: '較 Q1 大增 +5 ppts', options: { bold: true, fill: COLORS.greenBg, color: '065F46', align: 'center' } }
  ]
];

slide4.addTable(nodeRows, {
  x: 0.8, y: 1.95, w: 5.5,
  colW: [1.45, 0.9, 0.9, 0.9, 1.35],
  fontFace: FONT, fontSize: 10,
  border: { type: 'solid', color: COLORS.border, pt: 0.75 },
  valign: 'middle',
  margin: [5, 4, 5, 4]
});

slide4.addText(
  '• 2奈米 (N2) 里程碑：筆記本《The 2nm Era Arrives》預告 2026 進入 N2 時代，3Q26 單季即貢獻高達 8% 晶圓營收（約 US$37 億美元）。\n' +
  '• 3奈米 (N3) 擴張：由 1Q26 的 25% 大幅拉升至 3Q26 的 33%，受惠 AI 加速器與高階手機晶片全面轉進 N3P/N3E。',
  {
    x: 0.8, y: 5.45, w: 5.5, h: 1.25,
    fontFace: FONT, fontSize: 10, color: COLORS.textMain, lineSpacingMultiple: 1.25
  }
);

// Right Card: Platform Breakdown Table
slide4.addShape(pres.ShapeType.roundRect, {
  x: 6.83, y: 1.3, w: 5.9, h: 5.55,
  fill: { color: COLORS.white },
  line: { color: COLORS.border, width: 1 }
});
slide4.addText('💻 應用平台營收佔比與季成長 (By Platform)', {
  x: 7.03, y: 1.45, w: 5.5, h: 0.4,
  fontFace: FONT, fontSize: 14, bold: true, color: COLORS.navy
});

const platRows = [
  [
    { text: '應用平台', options: { bold: true, fill: COLORS.navy, color: COLORS.white } },
    { text: '1Q26 佔比\n(筆記本 PDF)', options: { bold: true, fill: COLORS.navy, color: COLORS.white, align: 'center' } },
    { text: '3Q26 佔比\n(最新法說)', options: { bold: true, fill: COLORS.navy, color: COLORS.white, align: 'center' } },
    { text: '3Q26 季增率\n(QoQ)', options: { bold: true, fill: COLORS.navy, color: COLORS.white, align: 'center' } },
    { text: '核心動能說明', options: { bold: true, fill: COLORS.navy, color: COLORS.white } }
  ],
  [
    { text: '高效能運算 (HPC)', options: { bold: true, fill: COLORS.greenBg, color: '065F46' } },
    { text: '61%', options: { fill: COLORS.white, align: 'center' } },
    { text: '68%', options: { bold: true, fill: COLORS.greenBg, color: '065F46', align: 'center' } },
    { text: '+22%', options: { bold: true, fill: COLORS.greenBg, color: '065F46', align: 'center' } },
    { text: 'AI GPU / ASIC 需求爆發', options: { fill: COLORS.white } }
  ],
  [
    { text: '智慧型手機 (Smartphone)', options: { bold: true, fill: 'F1F5F9' } },
    { text: '26%', options: { fill: 'F1F5F9', align: 'center' } },
    { text: '21%', options: { bold: true, fill: 'F1F5F9', align: 'center' } },
    { text: '+10%', options: { bold: true, fill: COLORS.blueBg, color: '0369A1', align: 'center' } },
    { text: '旗艦新機備貨，營收穩增', options: { fill: 'F1F5F9' } }
  ],
  [
    { text: '車用電子 (Automotive)', options: { bold: true, fill: COLORS.white } },
    { text: '4%', options: { fill: COLORS.white, align: 'center' } },
    { text: '4%', options: { fill: COLORS.white, align: 'center' } },
    { text: '+23%', options: { bold: true, fill: COLORS.greenBg, color: '065F46', align: 'center' } },
    { text: 'ADAS / 自駕晶片需求回溫', options: { fill: COLORS.white } }
  ],
  [
    { text: '物聯網 (IoT)', options: { bold: true, fill: 'F1F5F9' } },
    { text: '6%', options: { fill: 'F1F5F9', align: 'center' } },
    { text: '4%', options: { fill: 'F1F5F9', align: 'center' } },
    { text: '+5%', options: { fill: 'F1F5F9', align: 'center' } },
    { text: '邊緣裝置溫和復甦', options: { fill: 'F1F5F9' } }
  ],
  [
    { text: '消費性電子 (DCE) & 其他', options: { bold: true, fill: COLORS.white } },
    { text: '3%', options: { fill: COLORS.white, align: 'center' } },
    { text: '3%', options: { fill: COLORS.white, align: 'center' } },
    { text: 'DCE -12%', options: { fill: COLORS.white, align: 'center' } },
    { text: '佔整體營收比重極低', options: { fill: COLORS.white } }
  ]
];

slide4.addTable(platRows, {
  x: 7.03, y: 1.95, w: 5.5,
  colW: [1.5, 0.9, 0.9, 0.9, 1.3],
  fontFace: FONT, fontSize: 10,
  border: { type: 'solid', color: COLORS.border, pt: 0.75 },
  valign: 'middle',
  margin: [5, 4, 5, 4]
});

slide4.addText(
  '• HPC 結構性統治：由 1Q26 的 61% 一路擴張至 3Q26 的 68%，單季營收季增 +22%，成為推動台積電創高的絕對引擎。\n' +
  '• 手機業務實質成長：雖然手機佔比由 26% 降至 21%，但單季絕對金額仍季增 +10%，係因 HPC 成長斜率過陡導致相對比重變化。',
  {
    x: 7.03, y: 5.45, w: 5.5, h: 1.25,
    fontFace: FONT, fontSize: 10, color: COLORS.textMain, lineSpacingMultiple: 1.25
  }
);

// ==========================================================
// SLIDE 5: Full-Year Guidance Revisions & CapEx
// ==========================================================
const slide5 = pres.addSlide();
addHeader(
  slide5,
  '四、 2026 全年展望「三度上修」軌跡與資本支出比對',
  '從 NotebookLM 1月凱基報告、4月法說、7月法說到 10月 Q3 法說會的全年目標調升歷程',
  5
);

const revRows = [
  [
    { text: '評估時間點 / 資料來源', options: { bold: true, fill: COLORS.navy, color: COLORS.white } },
    { text: '2026 全年美元營收年增率預估', options: { bold: true, fill: COLORS.navy, color: COLORS.white, align: 'center' } },
    { text: '2026 全年資本支出 (CapEx)', options: { bold: true, fill: COLORS.navy, color: COLORS.white, align: 'center' } },
    { text: '管理階層對 AI 需求與獲利之核心定調', options: { bold: true, fill: COLORS.navy, color: COLORS.white } }
  ],
  [
    { text: '2026 年 1 月 (4Q25 法說)\n[筆記本：KGI 凱基報告]', options: { bold: true, fill: COLORS.white } },
    { text: '接近 30% (~30%)\n(上修 5年 CAGR 至 25%)', options: { fill: COLORS.white, align: 'center' } },
    { text: '520 ~ 560 億美元', options: { fill: COLORS.white, align: 'center' } },
    { text: 'AI 需求是真的；長期毛利率目標由 53% 上調至 56% 以上', options: { fill: COLORS.white } }
  ],
  [
    { text: '2026 年 4 月 (1Q26 法說)\n[筆記本：經濟日報 / 工商時報]', options: { bold: true, fill: 'F1F5F9' } },
    { text: '首度上修：超過 30% (>30%)', options: { bold: true, fill: COLORS.blueBg, color: '0369A1', align: 'center' } },
    { text: '520 ~ 560 億美元\n(預估將靠攏高標 560 億)', options: { fill: 'F1F5F9', align: 'center' } },
    { text: 'AI 需求依然「Tight (吃緊)」，正用盡所有方法擴產滿足客戶', options: { fill: 'F1F5F9' } }
  ],
  [
    { text: '2026 年 7 月 (2Q26 法說)\n[筆記本：七月營收分析筆記]', options: { bold: true, fill: COLORS.white } },
    { text: '二度上修：超過 40% (>40%)', options: { bold: true, fill: COLORS.blueBg, color: '0369A1', align: 'center' } },
    { text: '上修至 600 ~ 640 億美元', options: { bold: true, fill: COLORS.amberBg, color: 'B45309', align: 'center' } },
    { text: '加速執行擴產劇本，CoWoS 與 N2/N3 產能全面追趕需求', options: { fill: COLORS.white } }
  ],
  [
    { text: '2026 年 10 月 (3Q26 法說)\n[最新：2026/10/15 法說會]', options: { bold: true, fill: COLORS.greenBg, color: '065F46' } },
    { text: '🔥 三度上修：約 44% (~44%)', options: { bold: true, fill: COLORS.greenBg, color: '065F46', align: 'center' } },
    { text: '維持 600 ~ 640 億美元\n(預告 2026-28 遠高於前三年)', options: { bold: true, fill: COLORS.greenBg, color: '065F46', align: 'center' } },
    { text: '魏哲家：「AI 需求極度強勁且遠大於供給，台積電獲利能力沒有天花板」', options: { bold: true, fill: COLORS.greenBg, color: '065F46' } }
  ]
];

slide5.addTable(revRows, {
  x: 0.6, y: 1.35, w: 12.13,
  colW: [2.6, 2.5, 2.5, 4.53],
  fontFace: FONT, fontSize: 10.5,
  border: { type: 'solid', color: COLORS.border, pt: 0.75 },
  valign: 'middle',
  margin: [7, 6, 7, 6]
});

// Highlight Box at bottom of Slide 5
slide5.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 5.15, w: 12.13, h: 1.7,
  fill: { color: COLORS.navy }
});
slide5.addText('🎯 全年目標上修與資本支出關鍵洞察', {
  x: 0.85, y: 5.28, w: 11.6, h: 0.35,
  fontFace: FONT, fontSize: 14, bold: true, color: COLORS.gold
});
slide5.addText(
  '1. 全年營收成長率上修幅度高達 +14 個百分點：從年初預估「接近 30%」一路三度上修至「約 44%」，反映 AI 晶片與先進封裝產能開出速度優於預期。\n' +
  '2. 資本支出較年初調升約 80 億美元 (+15%)：由 Q1 的 520~560 億美元調升至 600~640 億美元，並確認 2026–2028 三年資本支出將顯著超越過去三年的 1,010 億美元。\n' +
  '3. 定價與毛利率防線：儘管海外廠（美、日）與 N2 初期量產帶來 2%~4% 毛利率稀釋，但透過技術價值定價與極高稼動率，全年毛利率穩守 66% 高檔。',
  {
    x: 0.85, y: 5.68, w: 11.6, h: 1.05,
    fontFace: FONT, fontSize: 11, color: COLORS.white, lineSpacingMultiple: 1.3
  }
);

// ==========================================================
// SLIDE 6: 4Q26 Guidance & Strategic Takeaways
// ==========================================================
const slide6 = pres.addSlide();
addHeader(
  slide6,
  '五、 2026 第四季 (4Q26) 最新財測指引與策略總結',
  '台積電 10/15 法說會釋出之 4Q26 營運展望：單季營收挑戰 533 億美元（約 1.7 兆新台幣）續創新高',
  6
);

// 4Q26 Guidance KPI Cards (Top Row)
const q4Cards = [
  { title: '4Q26 美元營收財測', val: '523 ~ 543 億美元', sub: '中位數 533 億美元\n季增 +14.1% / 年增 +58%', border: COLORS.accent },
  { title: '4Q26 台幣營收折算', val: '1.66兆 ~ 1.73兆元', sub: '匯率假設 1 USD = 31.8 NTD\n中位數約 1 兆 6,949 億元', border: COLORS.primary },
  { title: '4Q26 毛利率財測', val: '65.0% ~ 67.0%', sub: '中位數 66.0%\n維持 65%+ 超高獲利結構', border: COLORS.green },
  { title: '4Q26 營業利益率財測', val: '56.5% ~ 58.5%', sub: '中位數 57.5%\n連四季維持 56%+ 高水準', border: COLORS.gold }
];

q4Cards.forEach((c, i) => {
  const xPos = 0.6 + i * 3.1;
  slide6.addShape(pres.ShapeType.roundRect, {
    x: xPos, y: 1.35, w: 2.83, h: 1.9,
    fill: { color: COLORS.white },
    line: { color: c.border, width: 1.5 }
  });
  slide6.addShape(pres.ShapeType.rect, {
    x: xPos, y: 1.35, w: 2.83, h: 0.08,
    fill: { color: c.border }
  });
  slide6.addText(c.title, {
    x: xPos + 0.15, y: 1.5, w: 2.53, h: 0.3,
    fontFace: FONT, fontSize: 11.5, bold: true, color: COLORS.textMuted
  });
  slide6.addText(c.val, {
    x: xPos + 0.15, y: 1.85, w: 2.53, h: 0.5,
    fontFace: FONT, fontSize: 19, bold: true, color: COLORS.navy
  });
  slide6.addText(c.sub, {
    x: xPos + 0.15, y: 2.4, w: 2.53, h: 0.7,
    fontFace: FONT, fontSize: 10, color: COLORS.textMain, lineSpacingMultiple: 1.2
  });
});

// Strategic Takeaways Box (Bottom)
slide6.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 3.5, w: 12.13, h: 3.35,
  fill: { color: COLORS.white },
  line: { color: COLORS.border, width: 1 }
});
slide6.addText('📌 綜合結論：NotebookLM 筆記本預期 vs. 3Q26 實際成果驗證', {
  x: 0.85, y: 3.65, w: 11.6, h: 0.4,
  fontFace: FONT, fontSize: 15, bold: true, color: COLORS.navy
});

const summaryPoints = [
  '1. 【指標達成率滿分】無論是 2Q26 或 3Q26，台積電實際美元營收（Q2 達成率 101.5%、Q3 達成率 103.4%）與營業利益率（Q2 達成率 104.9%、Q3 達成率 103.0%）全面超越財測中位數並突破財測高標，驗證筆記本中「7月營收奠定挑戰高標基礎」的預判完全準確。',
  '2. 【2奈米 (N2) 商業化大成功】筆記本 Q1 資料中 N2 尚處於量產前準備，到了 3Q26 單季營收佔比已迅速衝上 8%（加計 3nm 33%、5nm 28%，先進製程佔比高達 79%），成為全球唯一能大規模量產 2 奈米與供應 AI 晶片的晶圓代工廠。',
  '3. 【AI 與 HPC 需求無天花板】HPC 佔比由 Q1 的 61% 擴大至 Q3 的 68%（單季季增 22%）。董事長魏哲家直指 AI 需求極度強勁且遠大於供給，全年美元營收成長率三度上修至 ~44%，4Q26 單季營收更將首度突破 523~543 億美元大關。',
  '4. 【中長期獲利預估持續上調】對照筆記本內 1 月份凱基報告（當時預估全年成長近 30%、CapEx 520-560 億美元），目前實際基本面已全面超越年初外資最樂觀情境。'
];

slide6.addText(summaryPoints.join('\n\n'), {
  x: 0.85, y: 4.15, w: 11.6, h: 2.5,
  fontFace: FONT, fontSize: 11, color: COLORS.textMain, lineSpacingMultiple: 1.25
});

const outPath1 = path.join('C:\\Users\\cr_hu\\OneDrive\\Desktop', 'TSMC_2026_Q3_法說會實績與指標達成率比對報告.pptx');
const outPath2 = path.join(__dirname, 'TSMC_2026_Q3_法說會實績與指標達成率比對報告.pptx');

Promise.all([
  pres.writeFile({ fileName: outPath1 }),
  pres.writeFile({ fileName: outPath2 })
]).then(() => {
  console.log('Successfully generated PPTX at:');
  console.log('1.', outPath1);
  console.log('2.', outPath2);
}).catch(err => {
  console.error('Error generating PPTX:', err);
  process.exit(1);
});
