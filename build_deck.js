/**
 * Build instructions:
 * 1) npm init -y
 * 2) npm i pptxgenjs
 * 3) node build_deck.js
 *
 * Output:
 * /output/nuriy_compliance_infrastructure_deck.pptx
 */

const fs = require('fs');
const PptxGenJS = require('pptxgenjs');

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5 (16:9)
pptx.author = 'Nuriy';
pptx.company = 'Nuriy';
pptx.subject = 'Investor deck';
pptx.title = 'Nuriy Compliance Infrastructure Deck';
pptx.lang = 'en-US';

const COLORS = {
  deepGreen: '1F4D3A',
  mutedGold: 'B08D57',
  softIvory: 'F7F3EA',
  clayBeige: 'E9DECC',
  charcoal: '2F2F2F',
  white: 'FFFFFF',
  lineLight: 'CFC6B7',
};

const FONTS = {
  heading: 'Georgia',
  body: 'Calibri',
};

const PAGE = {
  w: 13.333,
  h: 7.5,
  marginX: 0.65,
  titleY: 0.45,
  bodyY: 1.45,
  bodyW: 12.0,
  footerY: 7.1,
};

function addBase(slide, slideNum, opts = {}) {
  slide.background = { color: COLORS.softIvory };

  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: PAGE.w,
    h: 0.14,
    fill: { color: COLORS.deepGreen },
    line: { color: COLORS.deepGreen },
  });

  slide.addShape(pptx.ShapeType.line, {
    x: PAGE.marginX,
    y: PAGE.footerY - 0.08,
    w: PAGE.w - PAGE.marginX * 2,
    h: 0,
    line: { color: COLORS.lineLight, pt: 1 },
  });

  slide.addText('nuriy', {
    x: PAGE.marginX,
    y: PAGE.footerY,
    w: 1.5,
    h: 0.22,
    fontFace: FONTS.body,
    color: COLORS.deepGreen,
    fontSize: 11,
    bold: true,
  });

  if (opts.footerLeft) {
    slide.addText(opts.footerLeft, {
      x: PAGE.marginX + 1.2,
      y: PAGE.footerY,
      w: 3.2,
      h: 0.22,
      fontFace: FONTS.body,
      color: COLORS.charcoal,
      fontSize: 10,
    });
  }

  slide.addText(String(slideNum), {
    x: PAGE.w - PAGE.marginX - 0.4,
    y: PAGE.footerY,
    w: 0.4,
    h: 0.22,
    fontFace: FONTS.body,
    color: COLORS.charcoal,
    fontSize: 10,
    align: 'right',
  });
}

function addTitle(slide, text, subtitle) {
  slide.addText(text, {
    x: PAGE.marginX,
    y: PAGE.titleY,
    w: 11.6,
    h: 0.62,
    fontFace: FONTS.heading,
    color: COLORS.deepGreen,
    bold: true,
    fontSize: 40,
  });

  if (subtitle) {
    slide.addText(subtitle, {
      x: PAGE.marginX,
      y: PAGE.titleY + 0.72,
      w: 10.6,
      h: 0.36,
      fontFace: FONTS.body,
      color: COLORS.charcoal,
      fontSize: 20,
    });
  }
}

function addHeadline(slide, text) {
  slide.addText(text, {
    x: PAGE.marginX,
    y: PAGE.titleY,
    w: 11.8,
    h: 0.62,
    fontFace: FONTS.heading,
    color: COLORS.deepGreen,
    bold: true,
    fontSize: 33,
  });
}

function addBullets(slide, bullets, opts = {}) {
  const x = opts.x ?? PAGE.marginX;
  const y = opts.y ?? PAGE.bodyY;
  const w = opts.w ?? 6.1;
  const h = opts.h ?? 3.9;
  const fontSize = opts.fontSize ?? 20;

  const runs = bullets.map((b) => ({
    text: b,
    options: { bullet: { indent: fontSize * 0.9 } },
  }));

  slide.addText(runs, {
    x,
    y,
    w,
    h,
    fontFace: FONTS.body,
    color: COLORS.charcoal,
    fontSize,
    paraSpaceAfterPt: 11,
    breakLine: true,
  });
}

function addNotes(slide, lines) {
  slide.addNotes(lines.map((l) => `* ${l}`));
}

// Slide 1: Cover
{
  const slide = pptx.addSlide();
  addBase(slide, 1, { footerLeft: 'hello@nuriy.com, nuriy.com' });
  addTitle(slide, 'Nuriy', 'Verification Infrastructure for Jewelry Sustainability');

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 8.75,
    y: 2.1,
    w: 3.7,
    h: 2.4,
    rectRadius: 0.08,
    fill: { color: COLORS.clayBeige },
    line: { color: COLORS.lineLight, pt: 1 },
    shadow: { type: 'outer', color: '999999', blur: 1, angle: 45, distance: 1, opacity: 0.15 },
  });
  slide.addText('Mandatory\ncompliance\ninfrastructure', {
    x: 9.05,
    y: 2.55,
    w: 3.1,
    h: 1.7,
    fontFace: FONTS.heading,
    color: COLORS.deepGreen,
    fontSize: 23,
    align: 'center',
    valign: 'mid',
    bold: true,
  });

  addNotes(slide, [
    'Nuriy is building the independent verification infrastructure layer for jewelry sustainability.',
    'The core thesis is that compliance is becoming mandatory, spend becomes non discretionary.',
    'We are not leading with marketplace dynamics, we are leading with legal and procurement requirements.',
  ]);
}

// Slide 2: Problem
{
  const slide = pptx.addSlide();
  addBase(slide, 2);
  addHeadline(slide, 'Every Jewelry Brand Faces a Compliance Deadline');
  addBullets(slide, [
    'Regulatory mandate, EU CSDDD due diligence by 2027',
    'Legal exposure from unverified claims and supply chain opacity',
    'Infrastructure void, no independent product level verification standard',
  ]);

  const tx = 7.2;
  const ty = 2.2;
  const tw = 5.2;

  slide.addShape(pptx.ShapeType.line, {
    x: tx,
    y: ty + 1.7,
    w: tw,
    h: 0,
    line: { color: COLORS.deepGreen, pt: 2.2 },
  });

  const points = [
    { year: '2024', label: 'Signal\nbuild', x: tx + 0.15 },
    { year: '2025', label: 'Procurement\nintegration', x: tx + 1.65 },
    { year: '2026', label: 'Audit\nscale', x: tx + 3.05 },
    { year: '2027', label: 'CSDDD\ndeadline', x: tx + 4.4 },
  ];

  points.forEach((p) => {
    slide.addShape(pptx.ShapeType.ellipse, {
      x: p.x,
      y: ty + 1.57,
      w: 0.22,
      h: 0.22,
      fill: { color: COLORS.mutedGold },
      line: { color: COLORS.mutedGold },
    });
    slide.addText(p.year, {
      x: p.x - 0.25,
      y: ty + 1.87,
      w: 0.75,
      h: 0.2,
      fontFace: FONTS.body,
      fontSize: 11,
      color: COLORS.deepGreen,
      align: 'center',
      bold: true,
    });
    slide.addText(p.label, {
      x: p.x - 0.5,
      y: ty + 0.5,
      w: 1.3,
      h: 0.8,
      fontFace: FONTS.body,
      fontSize: 10,
      color: COLORS.charcoal,
      align: 'center',
    });
  });

  addNotes(slide, [
    'The problem is timing and liability, not awareness.',
    'Every major brand will need evidence that claims and sourcing data are defensible.',
    'Without product level standards, legal teams cannot operationalize compliance at scale.',
  ]);
}

// Slide 3: Market Opportunity
{
  const slide = pptx.addSlide();
  addBase(slide, 3);
  addHeadline(slide, 'A $300B+ Industry Requires Third-Party Verification');
  addBullets(slide, [
    'Global jewelry market exceeds $300B',
    'Compliance spend is emerging and unavoidable',
    'Supply chain transparency software is expanding quickly (illustrative)',
  ], { w: 6.2 });

  slide.addText('Jewelry is the beachhead, expansion follows.', {
    x: PAGE.marginX,
    y: 5.55,
    w: 6.2,
    h: 0.35,
    fontFace: FONTS.body,
    fontSize: 16,
    color: COLORS.deepGreen,
    italic: true,
  });

  const cx = 8.4;
  const cy = 2.0;
  const bars = [
    { label: 'Jewelry\n$300B+', h: 2.8, color: COLORS.deepGreen },
    { label: 'Compliance\nspend', h: 2.2, color: COLORS.mutedGold },
    { label: 'Supply Chain\nSoftware', h: 1.7, color: '6A8E7A' },
  ];

  bars.forEach((b, i) => {
    const x = cx + i * 1.45;
    const y = cy + (3.0 - b.h);
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y,
      w: 1.1,
      h: b.h,
      rectRadius: 0.06,
      fill: { color: b.color },
      line: { color: b.color },
    });
    slide.addText(b.label, {
      x: x - 0.2,
      y: cy + 3.12,
      w: 1.5,
      h: 0.6,
      fontFace: FONTS.body,
      fontSize: 10,
      color: COLORS.charcoal,
      align: 'center',
    });
  });

  addNotes(slide, [
    'This is a large base category where verification has been underbuilt.',
    'Budgets shift from discretionary brand spend into mandatory compliance operations.',
    'We start where risk is high, then replicate the infrastructure model into adjacent categories.',
  ]);
}

// Slide 4: LEED analogy
{
  const slide = pptx.addSlide();
  addBase(slide, 4);
  addHeadline(slide, 'What LEED Did for Buildings, Nuriy Does for Jewelry');

  const colY = 1.55;
  const boxH = 0.62;
  const leftX = 0.9;
  const rightX = 7.1;

  const flowLeft = ['Earn certification', 'Display badge', 'Market expects it', 'Standard emerges'];
  const flowRight = ['Earn score', 'Display label', 'Buyers expect it', 'Standard emerges'];

  [flowLeft, flowRight].forEach((flow, side) => {
    const startX = side === 0 ? leftX : rightX;
    const accent = side === 0 ? '8C9C91' : COLORS.deepGreen;
    slide.addShape(pptx.ShapeType.rect, {
      x: startX,
      y: 1.18,
      w: 5.3,
      h: 0.35,
      fill: { color: side === 0 ? 'DDE4DD' : 'D8E6DF' },
      line: { color: 'DDE4DD' },
    });
    slide.addText(side === 0 ? 'LEED Pattern' : 'Nuriy Pattern', {
      x: startX + 0.12,
      y: 1.22,
      w: 2.2,
      h: 0.22,
      fontFace: FONTS.body,
      fontSize: 12,
      color: COLORS.charcoal,
      bold: true,
    });

    flow.forEach((item, idx) => {
      const y = colY + idx * 0.8;
      slide.addShape(pptx.ShapeType.roundRect, {
        x: startX,
        y,
        w: 4.6,
        h: boxH,
        rectRadius: 0.05,
        fill: { color: COLORS.white },
        line: { color: accent, pt: 1.2 },
      });
      slide.addText(item, {
        x: startX + 0.18,
        y: y + 0.16,
        w: 4.2,
        h: 0.28,
        fontFace: FONTS.body,
        color: COLORS.charcoal,
        fontSize: 15,
      });
      if (idx < flow.length - 1) {
        slide.addShape(pptx.ShapeType.chevron, {
          x: startX + 4.75,
          y: y + 0.17,
          w: 0.42,
          h: 0.28,
          fill: { color: accent },
          line: { color: accent },
        });
      }
    });
  });

  addBullets(slide, [
    'Scoring engine, 0 to 100 across 5 dimensions',
    'Certification standard and licensing',
    'Infrastructure layer, APIs and audit ledger',
  ], { x: 0.9, y: 5.05, w: 11.5, h: 1.5, fontSize: 16 });

  addNotes(slide, [
    'The LEED analogy clarifies behavior change and market expectation formation.',
    'Nuriy turns fragmented evidence into a consistent product level score and label.',
    'As adoption scales, the scoring framework becomes the operating standard for procurement.',
  ]);
}

// Slide 5: Product Architecture
{
  const slide = pptx.addSlide();
  addBase(slide, 5);
  addHeadline(slide, 'Proprietary Verification Infrastructure');

  const layers = [
    'Nuriy Score Methodology (patent pending)',
    'Audit Synthesis Pipeline (AI reasoning and evidence extraction)',
    'Evidence Classification (Tier A audited, Tier B documented, Tier C missing)',
    'Integrity Layer (cryptographic hashing, tamper evident records)',
  ];

  const baseX = 1.0;
  const topY = 1.6;
  layers.forEach((layer, idx) => {
    const y = topY + idx * 1.05;
    slide.addShape(pptx.ShapeType.roundRect, {
      x: baseX + idx * 0.35,
      y,
      w: 8.8 - idx * 0.7,
      h: 0.78,
      rectRadius: 0.06,
      fill: { color: idx % 2 === 0 ? 'E3ECE7' : 'D6E4DC' },
      line: { color: COLORS.deepGreen, pt: 1 },
    });
    slide.addText(layer, {
      x: baseX + idx * 0.35 + 0.22,
      y: y + 0.23,
      w: 8.2 - idx * 0.7,
      h: 0.32,
      fontFace: FONTS.body,
      fontSize: 14,
      color: COLORS.charcoal,
    });
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 9.8,
    y: 2.45,
    w: 2.45,
    h: 1.55,
    rectRadius: 0.06,
    fill: { color: COLORS.clayBeige },
    line: { color: COLORS.mutedGold, pt: 1.2 },
  });
  slide.addText('Data moat\ngrows with\nevery audit.', {
    x: 10.0,
    y: 2.77,
    w: 2.05,
    h: 1.0,
    fontFace: FONTS.heading,
    fontSize: 18,
    color: COLORS.deepGreen,
    align: 'center',
    bold: true,
  });

  addNotes(slide, [
    'The stack is purpose built for verifiable outcomes, not generic ESG reporting.',
    'Each completed audit expands training and benchmark quality for future scoring.',
    'Integrity controls make records tamper evident for legal and procurement workflows.',
  ]);
}

// Slide 6: Regulatory Tailwinds
{
  const slide = pptx.addSlide();
  addBase(slide, 6);
  addHeadline(slide, 'Mandatory Compliance Creates Non-Discretionary Demand');

  addBullets(slide, [
    'EU CSDDD (2027)',
    'UK Modern Slavery Act',
    'US pending supply chain obligations (emerging)',
    'EU taxonomy and ESG claim substantiation',
  ], { w: 6.4 });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 7.35,
    y: 2.0,
    w: 5.1,
    h: 2.3,
    rectRadius: 0.06,
    fill: { color: 'EFE7D8' },
    line: { color: COLORS.mutedGold, pt: 1 },
  });
  slide.addText('"Brands buy because legal and procurement require proof."', {
    x: 7.65,
    y: 2.55,
    w: 4.5,
    h: 1.15,
    fontFace: FONTS.heading,
    fontSize: 24,
    italic: true,
    color: COLORS.deepGreen,
    valign: 'mid',
  });

  addNotes(slide, [
    'The buying trigger is institutional risk management.',
    'As obligations tighten, verification moves from optional to required infrastructure.',
    'This is why Nuriy can price as compliance software and certification infrastructure.',
  ]);
}

// Slide 7: Business Model
{
  const slide = pptx.addSlide();
  addBase(slide, 7);
  addHeadline(slide, 'Compliance-Driven B2B SaaS + Certification Licensing');

  const priorities = [
    '1) B2B SaaS subscriptions ($299 to $999 per month)',
    '2) Enterprise contracts ($50K+ annual)',
    '3) Certification licensing (earn and display label)',
    '4) API verification fees (per check at volume)',
    '5) Consumer features (secondary)',
  ];
  addBullets(slide, priorities, { x: 0.9, y: 1.5, w: 6.7, h: 3.7, fontSize: 16 });

  const tableX = 7.6;
  const tableY = 1.7;
  const colW = [1.2, 1.55, 1.2, 1.6];
  const rows = [
    ['Tier', 'Price (mo)', 'Score cap', 'Support'],
    ['Basic', '$299', '50', 'Email'],
    ['Standard', '$599', '200', 'Priority email'],
    ['Pro', '$999', '500', 'Dedicated CSM'],
    ['Enterprise', 'Custom', 'Unlimited', 'Procurement + API'],
  ];

  rows.forEach((row, r) => {
    let x = tableX;
    row.forEach((cell, c) => {
      const isHeader = r === 0;
      slide.addShape(pptx.ShapeType.rect, {
        x,
        y: tableY + r * 0.53,
        w: colW[c],
        h: 0.53,
        fill: { color: isHeader ? COLORS.deepGreen : COLORS.white },
        line: { color: COLORS.lineLight, pt: 0.8 },
      });
      slide.addText(cell, {
        x: x + 0.06,
        y: tableY + r * 0.53 + 0.14,
        w: colW[c] - 0.12,
        h: 0.2,
        fontFace: FONTS.body,
        fontSize: 10.5,
        color: isHeader ? COLORS.white : COLORS.charcoal,
        bold: isHeader,
        align: c === 0 ? 'left' : 'center',
      });
      x += colW[c];
    });
  });

  addNotes(slide, [
    'The revenue model prioritizes recurring B2B infrastructure contracts first.',
    'Certification licensing and API checks add scalable upside as standards adoption increases.',
    'Consumer functionality remains secondary and supports trust signaling.',
  ]);
}

// Slide 8: GTM
{
  const slide = pptx.addSlide();
  addBase(slide, 8);
  addHeadline(slide, 'Three-Phase Playbook');

  const phases = [
    { t: 'Phase 1, Calibration (Now)', d: 'Stabilize scoring, design partners' },
    { t: 'Phase 2, Activation', d: 'URL in, score out tool, first paying B2B customers' },
    { t: 'Phase 3, Standard', d: 'Chrome extension, enterprise compliance sales, label program' },
  ];

  phases.forEach((p, i) => {
    const x = 1.0 + i * 4.1;
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y: 2.2,
      w: 3.5,
      h: 2.0,
      rectRadius: 0.06,
      fill: { color: i === 1 ? 'E5D7BE' : 'E2EBE6' },
      line: { color: COLORS.deepGreen, pt: 1 },
    });
    slide.addText(p.t, {
      x: x + 0.18,
      y: 2.45,
      w: 3.15,
      h: 0.45,
      fontFace: FONTS.heading,
      fontSize: 17,
      bold: true,
      color: COLORS.deepGreen,
    });
    slide.addText(p.d, {
      x: x + 0.18,
      y: 2.95,
      w: 3.1,
      h: 1.0,
      fontFace: FONTS.body,
      fontSize: 13,
      color: COLORS.charcoal,
    });

    if (i < 2) {
      slide.addShape(pptx.ShapeType.chevron, {
        x: x + 3.62,
        y: 2.95,
        w: 0.32,
        h: 0.45,
        fill: { color: COLORS.mutedGold },
        line: { color: COLORS.mutedGold },
      });
    }
  });

  addNotes(slide, [
    'Phase sequencing balances technical calibration and revenue activation.',
    'Early design partners provide evidence quality and procurement insight.',
    'Standardization tools then widen distribution and enterprise adoption velocity.',
  ]);
}

// Slide 9: Traction
{
  const slide = pptx.addSlide();
  addBase(slide, 9);
  addHeadline(slide, 'Infrastructure Ready, Demand Validated');

  const metrics = [
    ['13,588', 'organic signups'],
    ['752+', 'products scored'],
    ['12', 'verified jewelers (design partners)'],
    ['7,877', 'Instagram followers'],
  ];

  metrics.forEach((m, i) => {
    const x = 0.95 + (i % 2) * 6.1;
    const y = 1.75 + Math.floor(i / 2) * 1.55;
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y,
      w: 5.4,
      h: 1.18,
      rectRadius: 0.06,
      fill: { color: 'FFFFFF' },
      line: { color: COLORS.lineLight, pt: 1 },
      shadow: { type: 'outer', color: 'AAAAAA', blur: 1, angle: 45, distance: 1, opacity: 0.12 },
    });
    slide.addText(m[0], {
      x: x + 0.25,
      y: y + 0.2,
      w: 1.75,
      h: 0.45,
      fontFace: FONTS.heading,
      fontSize: 30,
      color: COLORS.deepGreen,
      bold: true,
    });
    slide.addText(m[1], {
      x: x + 2.0,
      y: y + 0.35,
      w: 3.1,
      h: 0.4,
      fontFace: FONTS.body,
      fontSize: 14,
      color: COLORS.charcoal,
    });
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.95,
    y: 5.1,
    w: 11.5,
    h: 0.72,
    rectRadius: 0.06,
    fill: { color: 'D8E6DF' },
    line: { color: COLORS.deepGreen, pt: 1 },
  });
  slide.addText('Milestone, April 1 public URL in score out launch', {
    x: 1.2,
    y: 5.33,
    w: 10.8,
    h: 0.3,
    fontFace: FONTS.body,
    fontSize: 15,
    color: COLORS.deepGreen,
    bold: true,
    align: 'center',
  });

  addNotes(slide, [
    'Traction indicates both audience pull and operational readiness.',
    'We already have scored products and verified design partners feeding calibration.',
    'Public launch milestone converts infrastructure into recurring customer acquisition motion.',
  ]);
}

// Slide 10: Competitive Landscape
{
  const slide = pptx.addSlide();
  addBase(slide, 10);
  addHeadline(slide, 'No Product-Level Verification Infrastructure Exists');

  const cols = [
    {
      title: 'What exists',
      items: [
        'Brand level ratings',
        'Cooperation dependent blockchain pilots',
        'Manual certifications',
        'Horizontal ESG platforms',
      ],
      color: 'E8DDD0',
    },
    {
      title: 'What Nuriy delivers',
      items: [
        'Product level scoring',
        'Works without brand cooperation',
        'AI scalable evidence synthesis',
        'Jewelry specific methodology',
      ],
      color: 'D9E8E1',
    },
  ];

  cols.forEach((c, idx) => {
    const x = 1.0 + idx * 6.2;
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y: 1.65,
      w: 5.3,
      h: 3.9,
      rectRadius: 0.06,
      fill: { color: c.color },
      line: { color: COLORS.lineLight, pt: 1 },
    });
    slide.addText(c.title, {
      x: x + 0.2,
      y: 1.88,
      w: 4.9,
      h: 0.35,
      fontFace: FONTS.heading,
      fontSize: 22,
      color: COLORS.deepGreen,
      bold: true,
    });
    addBullets(slide, c.items, { x: x + 0.2, y: 2.35, w: 4.8, h: 2.9, fontSize: 15 });
  });

  slide.addText('First product-level verification standard built for jewelry.', {
    x: 1.0,
    y: 5.75,
    w: 11.4,
    h: 0.36,
    fontFace: FONTS.body,
    fontSize: 17,
    color: COLORS.deepGreen,
    bold: true,
    align: 'center',
  });

  addNotes(slide, [
    'Most alternatives either stop at brand level or require heavy cooperation to function.',
    'Nuriy focuses on independent, product specific verification that can scale with AI workflows.',
    'This creates a durable gap in defensibility and enterprise utility.',
  ]);
}

// Slide 11: Vision
{
  const slide = pptx.addSlide();
  addBase(slide, 11);
  addHeadline(slide, 'Verification Infrastructure Across Consumer Goods');

  const roadmap = [
    ['Now', 'Jewelry'],
    ['2027', 'Fashion'],
    ['2027 to 2028', 'Cosmetics'],
    ['2028+', 'Food'],
  ];

  slide.addShape(pptx.ShapeType.line, {
    x: 1.2,
    y: 3.5,
    w: 10.9,
    h: 0,
    line: { color: COLORS.deepGreen, pt: 2 },
  });

  roadmap.forEach((r, i) => {
    const x = 1.45 + i * 2.8;
    slide.addShape(pptx.ShapeType.ellipse, {
      x,
      y: 3.35,
      w: 0.3,
      h: 0.3,
      fill: { color: i === 0 ? COLORS.mutedGold : COLORS.deepGreen },
      line: { color: i === 0 ? COLORS.mutedGold : COLORS.deepGreen },
    });
    slide.addText(r[0], {
      x: x - 0.4,
      y: 2.95,
      w: 1.1,
      h: 0.25,
      fontFace: FONTS.body,
      fontSize: 11,
      color: COLORS.charcoal,
      align: 'center',
      bold: true,
    });
    slide.addText(r[1], {
      x: x - 0.7,
      y: 3.75,
      w: 1.6,
      h: 0.28,
      fontFace: FONTS.heading,
      fontSize: 17,
      color: COLORS.deepGreen,
      align: 'center',
      bold: true,
    });
  });

  slide.addText('Jewelry is the wedge, infrastructure is the business.', {
    x: 1.0,
    y: 5.2,
    w: 11.4,
    h: 0.45,
    fontFace: FONTS.body,
    fontSize: 20,
    color: COLORS.deepGreen,
    italic: true,
    align: 'center',
  });

  addNotes(slide, [
    'Expansion follows the same compliance logic, high risk categories first.',
    'Jewelry gives a focused data set and clear category definition for standard formation.',
    'The long term asset is cross category verification infrastructure.',
  ]);
}

// Slide 12: Analogies
{
  const slide = pptx.addSlide();
  addBase(slide, 12);
  addHeadline(slide, 'Infrastructure Category Comparables');

  const tiles = [
    { n: 'Vital4', d: 'Compliance infrastructure' },
    { n: 'Visalaw.AI', d: 'Automation for regulatory workflows' },
    { n: 'FinQuery', d: 'Financial compliance and accounting infrastructure' },
  ];

  tiles.forEach((t, i) => {
    const x = 1.05 + i * 4.15;
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y: 2.0,
      w: 3.55,
      h: 2.2,
      rectRadius: 0.08,
      fill: { color: i === 1 ? 'E9DECC' : 'E2EBE6' },
      line: { color: COLORS.lineLight, pt: 1 },
    });
    slide.addText(t.n, {
      x: x + 0.2,
      y: 2.4,
      w: 3.1,
      h: 0.4,
      fontFace: FONTS.heading,
      fontSize: 24,
      color: COLORS.deepGreen,
      bold: true,
      align: 'center',
    });
    slide.addText(t.d, {
      x: x + 0.2,
      y: 2.95,
      w: 3.1,
      h: 0.8,
      fontFace: FONTS.body,
      fontSize: 13,
      color: COLORS.charcoal,
      align: 'center',
    });
  });

  addBullets(slide, [
    'Category analogies only, no claimed affiliations',
    'We are building the verification layer, not a storefront',
  ], { x: 2.2, y: 4.65, w: 9.3, h: 1.2, fontSize: 16 });

  addNotes(slide, [
    'These analogies help frame the category for investors who know compliance infrastructure.',
    'Nuriy follows a similar infrastructure pattern but with product level verification as the core unit.',
    'We are positioning as backbone software and standards, not commerce front end.',
  ]);
}

// Slide 13: Financial Model
{
  const slide = pptx.addSlide();
  addBase(slide, 13);
  addHeadline(slide, 'MRR Growth Through B2B SaaS');

  const targets = [
    ['Q1 2026', '$3K to $4K MRR'],
    ['Q2 2026', '$10K to $15K MRR'],
    ['H2 2026', '$40K to $60K MRR'],
  ];

  targets.forEach((t, i) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 1.0,
      y: 1.85 + i * 1.15,
      w: 6.9,
      h: 0.86,
      rectRadius: 0.05,
      fill: { color: i === 2 ? 'D8E6DF' : 'FFFFFF' },
      line: { color: COLORS.lineLight, pt: 1 },
    });
    slide.addText(t[0], {
      x: 1.25,
      y: 2.1 + i * 1.15,
      w: 1.8,
      h: 0.3,
      fontFace: FONTS.body,
      fontSize: 15,
      color: COLORS.charcoal,
      bold: true,
    });
    slide.addText(t[1], {
      x: 3.0,
      y: 2.05 + i * 1.15,
      w: 4.5,
      h: 0.4,
      fontFace: FONTS.heading,
      fontSize: 24,
      color: COLORS.deepGreen,
      bold: true,
    });
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 8.35,
    y: 2.15,
    w: 3.95,
    h: 2.6,
    rectRadius: 0.06,
    fill: { color: 'EFE7D8' },
    line: { color: COLORS.mutedGold, pt: 1.1 },
  });
  slide.addText('Unit economics\nplaceholders', {
    x: 8.6,
    y: 2.42,
    w: 3.4,
    h: 0.55,
    fontFace: FONTS.heading,
    fontSize: 20,
    color: COLORS.deepGreen,
    bold: true,
    align: 'center',
  });
  addBullets(slide, ['CAC, TBD', 'LTV, TBD', 'Burn, TBD'], {
    x: 8.8,
    y: 3.1,
    w: 3.0,
    h: 1.4,
    fontSize: 14,
  });

  addNotes(slide, [
    'Model focus is recurring MRR progression, not GMV metrics.',
    'Targets reflect ramp from initial SaaS subscriptions into enterprise contracts.',
    'Unit economics are placeholders and will tighten with live cohort data.',
  ]);
}

// Slide 14: Ask
{
  const slide = pptx.addSlide();
  addBase(slide, 14);
  addHeadline(slide, '$500K to Activate the Infrastructure');

  addBullets(slide, [
    'Integrity layer completion',
    'API build out and B2B sales execution',
    'Chrome extension preparation',
  ], { x: 1.0, y: 2.0, w: 6.4, h: 2.0, fontSize: 19 });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 7.4,
    y: 1.95,
    w: 4.9,
    h: 2.5,
    rectRadius: 0.07,
    fill: { color: 'D8E6DF' },
    line: { color: COLORS.deepGreen, pt: 1 },
  });
  slide.addText('Runway\n18 months', {
    x: 8.2,
    y: 2.45,
    w: 3.3,
    h: 1.1,
    fontFace: FONTS.heading,
    fontSize: 33,
    color: COLORS.deepGreen,
    bold: true,
    align: 'center',
  });

  slide.addText('Capital buys speed.', {
    x: 1.0,
    y: 5.25,
    w: 11.2,
    h: 0.45,
    fontFace: FONTS.heading,
    fontSize: 26,
    color: COLORS.deepGreen,
    italic: true,
    align: 'center',
  });

  addNotes(slide, [
    'This round funds the highest leverage infrastructure milestones.',
    'Primary uses are integrity completion, API scale, and enterprise go to market readiness.',
    'The goal is to convert readiness into contracted recurring revenue within runway.',
  ]);
}

// Slide 15: Closing
{
  const slide = pptx.addSlide();
  addBase(slide, 15, { footerLeft: 'hello@nuriy.com | nuriy.com' });
  addHeadline(slide, 'First-Mover in a Mandatory Compliance Market');
  addBullets(slide, [
    'Regulatory urgency, CSDDD 2027',
    'Proprietary infrastructure, patent pending, growing data moat',
    'Certification standard potential, LEED model',
  ], { x: 1.1, y: 2.2, w: 9.2, h: 2.6, fontSize: 22 });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 9.2,
    y: 2.0,
    w: 3.0,
    h: 2.8,
    rectRadius: 0.1,
    fill: { color: COLORS.clayBeige },
    line: { color: COLORS.lineLight, pt: 1 },
  });

  const ringCenterX = 10.7;
  const ringCenterY = 3.35;
  slide.addShape(pptx.ShapeType.donut, {
    x: ringCenterX - 0.72,
    y: ringCenterY - 0.72,
    w: 1.44,
    h: 1.44,
    fill: { color: COLORS.deepGreen },
    line: { color: COLORS.deepGreen, pt: 1 },
    rotate: 0,
    adjustPoint: 0.45,
  });
  slide.addText('Nuriy\nStandard', {
    x: 9.6,
    y: 4.0,
    w: 2.2,
    h: 0.7,
    fontFace: FONTS.body,
    fontSize: 13,
    color: COLORS.deepGreen,
    bold: true,
    align: 'center',
  });

  addNotes(slide, [
    'Nuriy enters at a moment when compliance obligations are becoming explicit and timed.',
    'Our edge is proprietary verification infrastructure plus a standard formation pathway.',
    'We are inviting investors into a category defining layer with strong enterprise pull.',
  ]);
}

fs.mkdirSync('/output', { recursive: true });
pptx.writeFile({ fileName: '/output/nuriy_compliance_infrastructure_deck.pptx' })
  .then(() => {
    console.log('Deck written to /output/nuriy_compliance_infrastructure_deck.pptx');
  })
  .catch((err) => {
    console.error('Failed to write deck:', err);
    process.exit(1);
  });
