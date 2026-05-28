const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, FootnoteReferenceRun,
  LevelFormat, ExternalHyperlink
} = require('docx');
const fs = require('fs-extra');
const config = require('./config');

// ─── Setup Output Directories ────────────────────────────────────────────────
const createOutputDirs = () => {
  fs.ensureDirSync(`${config.OUTPUT_DIR}/docs`);
  fs.ensureDirSync(`${config.OUTPUT_DIR}/pdfs`);
};

// ─── Helper Functions ──────────────────────────────────────────────────────────
const COLORS = config.COLORS;
const border = { style: BorderStyle.SINGLE, size: 4, color: COLORS.midGray };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
const FULL_W = config.FULL_W;

const h1 = (text, opts = {}) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  pageBreakBefore: opts.pageBreak !== false,
  children: [new TextRun({ text, font: 'Arial', size: 32, bold: true, color: COLORS.primary })]
});

const h2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 240, after: 120 },
  children: [new TextRun({ text, font: 'Arial', size: 26, bold: true, color: COLORS.primary })]
});

const h3 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  spacing: { before: 180, after: 80 },
  children: [new TextRun({ text, font: 'Arial', size: 24, bold: true, color: COLORS.secondary })]
});

const para = (text, opts = {}) => new Paragraph({
  spacing: { before: 80, after: 120 },
  alignment: opts.center ? AlignmentType.CENTER : AlignmentType.JUSTIFIED,
  children: [new TextRun({ text, font: 'Arial', size: 22, bold: opts.bold, italic: opts.italic, color: opts.color })]
});

const bullet = (text, level = 0) => new Paragraph({
  numbering: { reference: 'bullets', level },
  spacing: { before: 40, after: 40 },
  children: [new TextRun({ text, font: 'Arial', size: 22 })]
};

const blank = () => new Paragraph({ children: [new TextRun('')], spacing: { before: 80, after: 80 } });
const divider = () => new Paragraph({
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: COLORS.gold } },
  children: [new TextRun('')],
  spacing: { before: 160, after: 160 }
});

// ─── Table Helpers ──────────────────────────────────────────────────────────────
const makeCell = (text, opts = {}) => new TableCell({
  borders: opts.noBorder ? noBorders : borders,
  width: { size: opts.w || FULL_W / 2, type: WidthType.DXA },
  shading: { fill: opts.fill || COLORS.white, type: ShadingType.CLEAR },
  margins: { top: 100, bottom: 100, left: 120, right: 120 },
  verticalAlign: VerticalAlign.CENTER,
  columnSpan: opts.span,
  children: [new Paragraph({
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
    children: [new TextRun({ text: String(text || ''), font: 'Arial', size: opts.size || 20, bold: opts.bold, color: opts.color || '000000' })]
  })]
});

const headerRow = (cols, widths, fill = COLORS.primary) => new TableRow({
  tableHeader: true,
  children: cols.map((c, i) => new TableCell({
    borders,
    width: { size: widths[i], type: WidthType.DXA },
    shading: { fill, type: ShadingType.CLEAR },
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: c, font: 'Arial', size: 20, bold: true, color: COLORS.white })]
    })]
  }))
});

const dataRow = (cells, widths, fill = COLORS.white) => new TableRow({
  children: cells.map((c, i) => new TableCell({
    borders,
    width: { size: widths[i], type: WidthType.DXA },
    shading: { fill, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      children: [new TextRun({ text: String(c ?? ''), font: 'Arial', size: 20, color: '000000' })]
    })]
  }))
});

const altRow = (cells, widths, idx) => dataRow(cells, widths, idx % 2 === 0 ? COLORS.white : COLORS.lightGray);

// ─── FOOTNOTES DATA ────────────────────────────────────────────────────────────
const footnotes = {
  1:  { children: [new Paragraph({ children: [new TextRun({ text: 'Philippine Statistics Authority–RSSO BARMM. (2025). BARMM\'s economy expands by 2.7 percent in 2024.' })] })] },
  2:  { children: [new Paragraph({ children: [new TextRun({ text: 'Ministry of Finance and Budget Management–BARMM (MFBM). (2025). BARMM Provincial Economic Trends. MFBM Economics Division.' })] })] },
  3:  { children: [new Paragraph({ children: [new TextRun({ text: 'Silva, A. M. (2026). Reimagining BARMM\'s economic sectors: Building an integrated investment roadmap (2026–2036).' })] })] },
  4:  { children: [new Paragraph({ children: [new TextRun({ text: 'SunStar Publishing Inc. (2025, October 21). BARMM hits ₱5B in 2025 investment. SunStar Business.' })] })] },
  5:  { children: [new Paragraph({ children: [new TextRun({ text: 'Bangsamoro Planning and Development Authority (BPDA). (2023). 2nd Bangsamoro Development Plan 2023–2028.' })] })] },
  6:  { children: [new Paragraph({ children: [new TextRun({ text: 'Department of Energy. (2025). DOE spearheads comprehensive measures to resolve Basilan\'s energy challenges.' })] })] },
  7:  { children: [new Paragraph({ children: [new TextRun({ text: 'Luwaran. (2024, December 8). Lanao del Sur leads BARMM economic growth. https://www.luwaran.com' })] })] },
  8:  { children: [new Paragraph({ children: [new TextRun({ text: 'BIMP-EAGA. (2024). BIMP-EAGA Strategic Action Plan 2025–2030. ASEAN Secretariat.' })] })] },
  9:  { children: [new Paragraph({ children: [new TextRun({ text: 'Ministry of Trade, Investments and Tourism–BARMM (MTIT). (2026). Bangsamoro Revealing Its Hidden Treasure – Investment Roadmap 2026–2035.' })] })] },
  10: { children: [new Paragraph({ children: [new TextRun({ text: 'Silva, A. M. & Unsi, M. (2026). The Emerging Bangsamoro: A Hub for Resilient and Ethical Growth.' })] })] },
};

// ─── COVER SECTION ─────────────────────────────────────────────────────────────
function buildCoverSection() {
  return [
    blank(), blank(),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 800, after: 200 },
      children: [new TextRun({ text: 'BANGSAMORO AUTONOMOUS REGION IN MUSLIM MINDANAO', { size: 24, bold: true, color: COLORS.primary })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 100, after: 100 },
      children: [new TextRun({ text: 'MINISTRY OF TRADE, INVESTMENTS AND TOURISM (MTIT)', { size: 22, bold: true, color: COLORS.primary })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 100, after: 100 },
      children: [new TextRun({ text: 'BUREAU OF INVESTMENTS (BOI-MTIT)', { size: 22, color: COLORS.primary })]
    }),
    blank(),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 400, after: 200 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: COLORS.gold } },
      children: [new TextRun({ text: config.TITLE, { size: 40, bold: true, color: COLORS.primary })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 200 },
      children: [new TextRun({ text: '(BIRD) 2026–2035', { size: 40, bold: true, color: COLORS.gold })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 100, after: 400 },
      border: { top: { style: BorderStyle.SINGLE, size: 8, color: COLORS.gold } },
      children: [new TextRun({ text: config.SUBTITLE, { size: 26, italic: true, color: COLORS.secondary })]
    }),
    blank(), blank(),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun(`Prepared by the Bureau of Investments – MTIT${blank}, size: 22, bold: true`) ]}),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun(`In collaboration with BPDA, MAFAR, MENRE, MPW, BBOI, BEZA, and Provincial LGUs`, { size: 20, italic: true })] }),
    blank(),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun(`${config.CONTACT.city} | ${config.DATE}`, { size: 22 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun('Data Currency: PSA 2023–2025 | BARMM Official Sources', { size: 20, italic: true, color: '666666' })] }),
    blank(), blank(),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun('For Policymakers · Development and Investment Planners · Investors', { size: 22, italic: true, color: COLORS.secondary })] }),
  ];
}

// ─── EXECUTIVE SUMMARY SECTION ─────────────────────────────────────────────────
function buildExecutiveSummary() {
  return [
    h1('Executive Summary', { pageBreak: true }),
    para('The Bangsamoro Autonomous Region in Muslim Mindanao (BARMM) stands at a decisive crossroads in its development journey — transitioning from post-conflict recovery to an emerging hub for inclusive and ethical growth...'),
    h3('Key Strategic Pillars'),
    bullet('Halal Economy: Positioning BARMM as the Philippines\' premier halal production and certification hub'),
    bullet('Green Economy & Environment: Leveraging BIMP-EAGA EGL framework, carbon markets, and REDD+'),
    bullet('Agro-Industrial Transformation: Developing integrated value chains from production to export-ready processing'),
    bullet('Infrastructure Enablement: Closing critical gaps in energy, transport, digital connectivity, and logistics'),
    h3('Key Economic Indicators at a Glance'),
    new Table({
      width: { size: FULL_W, type: WidthType.DXA },
      columnWidths: [3500, 3000, 2860],
      rows: [
        headerRow(['Indicator', 'Latest Value', 'Source / Year'], [3500, 3000, 2860]),
        altRow(['BARMM GRDP', '₱299.5 billion', 'PSA, 2025'], [3500, 3000, 2860], 0),
        altRow(['GRDP Growth Rate (2024)', '2.7%', 'PSA, 2025'], [3500, 3000, 2860], 1),
        altRow(['Services Share', '42.0% (₱125.7B)', 'PSA, 2025'], [3500, 3000, 2860], 0),
        altRow(['AFF Share', '32.4% (₱97.2B)', 'PSA, 2025'], [3500, 3000, 2860], 1),
        altRow(['Industry Share', '25.6% (₱76.6B)', 'PSA, 2025'], [3500, 3000, 2860], 0),
        altRow(['Population (2024)', '5,691,583', 'PSA, 2025'], [3500, 3000, 2860], 1),
        altRow(['Poverty Incidence (H1 2023)', '34.8%', 'PSA, 2024'], [3500, 3000, 2860], 0),
        altRow(['Functional Literacy Rate', '59.3% (lowest in PH)', 'PSA FLEMMS, 2024'], [3500, 3000, 2860], 1),
      ]
    }),
    para('Table ES-1. BARMM Key Economic Indicators (2024–2026). Sources: PSA-BARMM (2025); BBOI; MFBM.', { italic: true }),
    h3('Implementation Framework'),
    para('The BIRD 2026–2035 is organized around three implementation phases: Foundation Building (2026–2028), Acceleration (2029–2032), and Consolidation and Global Integration (2033–2035).'),
  ];
}

// ─── ADDITIONAL SECTIONS (Chapters 1-12) ──────────────────────────────────────
// Note: For brevity, these are abbreviated. Use your original complete chapter functions.
// Full implementations should include SWOT analysis, BEIE framework, Balanced Scorecard, etc.

function buildChapter1() {
  return [h1('Chapter 1: Introduction'), h2('1.1 Background and Rationale'), para('BARMM was formally established on February 26, 2019...')]
}

function buildChapter10() {
  return [
    h1('Chapter 10: Implementation Roadmap'),
    h2('Phase 1: Foundation Building (2026–2028)'),
    new Table({
      width: { size: FULL_W, type: WidthType.DXA },
      columnWidths: [2800, 4360, 2200],
      rows: [
        headerRow(['Priority Action', 'Key Deliverables', 'Lead Agency'], [2800, 4360, 2200]),
        altRow(['Halal Governance', 'Operationalize BHB with OIC/SMIIC-aligned processes', 'BHB / MTIT'], [2800, 4360, 2200], 0),
        altRow(['Digital Facilitation', 'Digital business registration (1-day target)', 'BBOI / DICT'], [2800, 4360, 2200], 1),
        altRow(['Green Economy Framework', 'Forestry Code enacted; JMC No. 2026-01 operationalized', 'MENRE / Parliament'], [2800, 4360, 2200], 0),
      ]
    }),
  ];
}

// ─── BUILD REFERENCES SECTION ──────────────────────────────────────────────────
function buildReferences() {
  const ref = (text) => new Paragraph({
    spacing: { before: 60, after: 60 },
    indent: { left: 720, hanging: 720 },
    children: [new TextRun({ text, font: 'Arial', size: 20 })]
  });
  return [
    h1('References'),
    ref('Bangsamoro Ministry of Trade, Investments and Tourism (MTIT). (2026). Bangsamoro Investment Roadmap Development (BIRD) 2026–2035.'),
    ref('Bangsamoro Planning and Development Authority (BPDA). (2023). 2nd Bangsamoro Development Plan 2023–2028.'),
    ref('Philippine Statistics Authority–RSSO BARMM (PSA). (2025). BARMM\'s economy expands by 2.7 percent in 2024.'),
    ref('World Bank. (2025). Global Halal Market Outlook 2025–2030.'),
  ];
}

// ─── MAIN DOCUMENT CREATION ────────────────────────────────────────────────────
async function generateDocument() {
  console.log('[BUILD] Creating BIRD 2026–2035 Investment Roadmap...');
  
  try {
    await createOutputDirs();
    
    const doc = new Document({
      footnotes,
      numbering: {
        config: [{
          reference: 'bullets',
          levels: [{
            level: 0, format: LevelFormat.BULLET, text: '•',
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 540, hanging: 360 } } }
          }]
        }]
      },
      styles: {
        default: { document: { run: { font: 'Arial', size: 22 } } },
        paragraphStyles: [
          { id: 'Heading1', name: 'Heading 1', run: { size: 32, bold: true, font: 'Arial', color: COLORS.primary } },
          { id: 'Heading2', name: 'Heading 2', run: { size: 26, bold: true, font: 'Arial', color: COLORS.primary } },
          { id: 'Heading3', name: 'Heading 3', run: { size: 24, bold: true, font: 'Arial', color: COLORS.secondary } },
        ]
      },
      sections: [{
        properties: {
          page: {
            size: { width: 12240, height: config.PAGE_HEIGHT },
            margin: { top: config.MARGIN, right: config.MARGIN, bottom: config.MARGIN, left: config.MARGIN }
          }
        },
        headers: {
          default: new Header({
            children: [new Paragraph({
              alignment: AlignmentType.RIGHT,
              border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: COLORS.gold } },
              children: [
                new TextRun({ text: `${config.TITLE} | BOI-MTIT, BARMM`, font: 'Arial', size: 18, color: COLORS.primary })
              ]
            })]
          })
        },
        footers: {
          default: new Footer({
            children: [new Paragraph({
              alignment: AlignmentType.CENTER,
              border: { top: { style: BorderStyle.SINGLE, size: 4, color: COLORS.gold } },
              children: [
                new TextRun({ text: 'For Policymakers, Development Planners, and Investors | Confidential until officially released | Page ', font: 'Arial', size: 16, color: '666666' }),
                new PageNumber()
              ]
            })]
          })
        },
        children: [
          ...buildCoverSection(),
          new Paragraph({ children: [new PageBreak()] }),
          ...buildExecutiveSummary(),
          ...buildChapter1(),
          ...buildChapter10(),
          ...buildReferences(),
        ]
      }]
    });

    await Packer.toBuffer(doc).then(buffer => {
      fs.writeFileSync(config.DOCX_PATH, buffer);
      console.log(`[SUCCESS] Document saved to: ${config.DOCX_PATH}`);
    });

  } catch (error) {
    console.error('[ERROR] Document generation failed:', error.message);
    throw error;
  }
}

// ─── EXPORT FOR MODULE USE ─────────────────────────────────────────────────────
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { generateDocument };
}

// ─── AUTO-RUN IF EXECUTED DIRECTLY ────────────────────────────────────────────
if (require.main === module) {
  generateDocument().catch(err => {
    console.error('Fatal error:', err);
    process.exit(1);
  });
}
