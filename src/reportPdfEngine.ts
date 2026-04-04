import fontkit from '@pdf-lib/fontkit';
import { PDFDocument, PDFFont, PDFPage, StandardFonts, rgb } from 'pdf-lib';
import {
  getEmbeddedAppRegularOtf,
  getEmbeddedReportBoldOtf,
  getEmbeddedReportRegularOtf,
} from './monieziFonts';

export interface ExpenseSummaryRow {
  name: string;
  amount: number;
  sharePct: number;
  linked: number;
  count: number;
}

export interface MileageQuarterRow {
  quarter: string;
  trips: number;
  miles: number;
  deduction: number;
}

export interface TaxSummaryPdfData {
  taxYear: string;
  businessName: string;
  ownerName: string;
  generatedAtLabel: string;
  reportingPeriodLabel: string;
  totalIncome: number;
  totalExpenses: number;
  netProfit: number;
  totalMiles: number;
  mileageDeduction: number;
  mileageRate: number;
  expenseItemsCount: number;
  ledgerTransactions: number;
  linkedReceipts: number;
  expenseCategoriesCount: number;
  topExpenseCategoryName: string;
  topExpenseCategoryAmount: number;
  topExpenseCategorySharePct: number;
  receiptCoveragePct: number;
  reviewCoveragePct: number;
  mileageCompletionPct: number;
  reviewedExpenseCount: number;
  pendingReviewCount: number;
  completeMileageCount: number;
  itemsRequiringAttention: number;
  expenseRows: ExpenseSummaryRow[];
  quarterlyMileage: MileageQuarterRow[];
  hasMileageRows: boolean;
  attentionItems: string[];
  currencySymbol: string;
}

type FontSet = {
  body: PDFFont;
  bold: PDFFont;
  kicker: PDFFont;
};

const PAGE = {
  width: 595.28,
  height: 841.89,
  marginX: 40,
  marginTop: 44,
  marginBottom: 40,
};

const CONTENT_WIDTH = PAGE.width - PAGE.marginX * 2;
const SECTION_HEADER_HEIGHT = 70;
const SECTION_INSET = 18;
const TABLE_HEADER_HEIGHT = 28;
const FOOTER_Y = PAGE.marginBottom - 10;

const COLORS = {
  ink: rgb(0.07, 0.11, 0.2),
  inkSoft: rgb(0.38, 0.45, 0.56),
  blue: rgb(0.17, 0.39, 0.89),
  line: rgb(0.86, 0.89, 0.94),
  panel: rgb(0.97, 0.98, 1),
  panelStrong: rgb(0.93, 0.96, 1),
  panelBorder: rgb(0.84, 0.89, 0.95),
  green: rgb(0.16, 0.70, 0.36),
  greenSoft: rgb(0.89, 0.97, 0.91),
  red: rgb(0.88, 0.25, 0.25),
  redSoft: rgb(1, 0.93, 0.93),
  yellow: rgb(0.87, 0.60, 0.10),
  yellowSoft: rgb(1, 0.96, 0.84),
  white: rgb(1, 1, 1),
};

const sanitizePdfText = (value: unknown) => String(value ?? '')
  .normalize('NFKD')
  .replace(/[\u2010-\u2015]/g, '-')
  .replace(/[\u2022\u00B7]/g, '-')
  .replace(/[\u2018\u2019]/g, "'")
  .replace(/[\u201C\u201D]/g, '"')
  .replace(/\u2026/g, '...')
  .replace(/\u00A0/g, ' ')
  .replace(/[^\x20-\x7E]/g, ' ')
  .replace(/\s+/g, ' ')
  .trim();

const formatCurrency = (symbol: string, value: number) => `${symbol}${Number(value || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
const formatNumber = (value: number, decimals = 0) => Number(value || 0).toLocaleString(undefined, { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
const formatPercent = (value: number) => `${Number(value || 0).toLocaleString(undefined, { minimumFractionDigits: value % 1 === 0 ? 0 : 1, maximumFractionDigits: 1 })}%`;

const splitLines = (text: string, font: PDFFont, size: number, maxWidth: number) => {
  const safeText = sanitizePdfText(text);
  if (!safeText) return [''];
  const words = safeText.split(/\s+/).filter(Boolean);
  const lines: string[] = [];
  let current = words[0] || '';

  for (let i = 1; i < words.length; i += 1) {
    const candidate = `${current} ${words[i]}`;
    if (font.widthOfTextAtSize(candidate, size) <= maxWidth) {
      current = candidate;
    } else {
      lines.push(current);
      current = words[i];
    }
  }

  lines.push(current);
  return lines;
};

const textBlockHeight = (text: string, font: PDFFont, size: number, width: number, lineGap = 4) => {
  const lines = splitLines(text, font, size, width);
  return lines.length * size + Math.max(0, lines.length - 1) * lineGap;
};

const drawTextBlock = (
  page: PDFPage,
  text: string,
  x: number,
  yTop: number,
  width: number,
  font: PDFFont,
  size: number,
  color = COLORS.inkSoft,
  lineGap = 4,
) => {
  const lines = splitLines(text, font, size, width);
  const lineHeight = size + lineGap;
  let y = yTop - size;
  lines.forEach(line => {
    page.drawText(sanitizePdfText(line), { x, y, size, font, color });
    y -= lineHeight;
  });
  return y;
};

const drawRule = (page: PDFPage, y: number) => {
  page.drawLine({
    start: { x: PAGE.marginX, y },
    end: { x: PAGE.width - PAGE.marginX, y },
    thickness: 1,
    color: COLORS.line,
  });
};

const drawInfoBox = (
  page: PDFPage,
  x: number,
  yTop: number,
  width: number,
  label: string,
  value: string,
  fonts: FontSet,
  height = 46,
) => {
  page.drawRectangle({ x, y: yTop - height, width, height, color: COLORS.panelStrong, borderColor: COLORS.panelBorder, borderWidth: 1 });
  page.drawText(sanitizePdfText(label).toUpperCase(), {
    x: x + 12,
    y: yTop - 16,
    size: 7.2,
    font: fonts.bold,
    color: COLORS.inkSoft,
    characterSpacing: 1.2,
  });
  const valueLines = splitLines(value, fonts.bold, 10.5, width - 24).slice(0, 2);
  let lineY = yTop - 31;
  valueLines.forEach(line => {
    page.drawText(sanitizePdfText(line), { x: x + 12, y: lineY, size: 10.5, font: fonts.bold, color: COLORS.ink });
    lineY -= 12;
  });
};

const drawMetricCard = (
  page: PDFPage,
  x: number,
  yTop: number,
  width: number,
  height: number,
  label: string,
  value: string,
  note: string,
  fonts: FontSet,
) => {
  page.drawRectangle({ x, y: yTop - height, width, height, color: COLORS.white, borderColor: COLORS.panelBorder, borderWidth: 1 });
  page.drawText(sanitizePdfText(label).toUpperCase(), {
    x: x + 14,
    y: yTop - 18,
    size: 7.1,
    font: fonts.bold,
    color: COLORS.inkSoft,
    characterSpacing: 1.2,
  });
  page.drawText(sanitizePdfText(value), {
    x: x + 14,
    y: yTop - 43,
    size: 17,
    font: fonts.bold,
    color: COLORS.ink,
  });
  drawTextBlock(page, note, x + 14, yTop - 58, width - 28, fonts.body, 7, COLORS.inkSoft, 2.3);
};

const drawSectionCard = (
  page: PDFPage,
  x: number,
  yTop: number,
  width: number,
  height: number,
  sectionNo: string,
  title: string,
  subtitle: string,
  fonts: FontSet,
) => {
  page.drawRectangle({ x, y: yTop - height, width, height, color: COLORS.white, borderColor: COLORS.panelBorder, borderWidth: 1 });
  page.drawRectangle({ x, y: yTop - SECTION_HEADER_HEIGHT, width, height: SECTION_HEADER_HEIGHT, color: COLORS.panel });
  page.drawText(`SECTION ${sanitizePdfText(sectionNo)}`, {
    x: x + SECTION_INSET,
    y: yTop - 20,
    size: 8.1,
    font: fonts.bold,
    color: COLORS.blue,
    characterSpacing: 1.4,
  });
  page.drawText(sanitizePdfText(title), {
    x: x + SECTION_INSET,
    y: yTop - 42,
    size: 14,
    font: fonts.bold,
    color: COLORS.ink,
  });
  const subtitleLines = splitLines(subtitle, fonts.body, 7.2, width - SECTION_INSET * 2).slice(0, 2);
  let subtitleY = yTop - 56;
  subtitleLines.forEach(line => {
    page.drawText(sanitizePdfText(line), { x: x + SECTION_INSET, y: subtitleY, size: 7.2, font: fonts.body, color: COLORS.inkSoft });
    subtitleY -= 10;
  });
  page.drawLine({ start: { x, y: yTop - SECTION_HEADER_HEIGHT }, end: { x: x + width, y: yTop - SECTION_HEADER_HEIGHT }, thickness: 1, color: COLORS.line });
  return yTop - SECTION_HEADER_HEIGHT - SECTION_INSET;
};

const progressTone = (value: number) => {
  if (value >= 90) return { track: COLORS.greenSoft, bar: COLORS.green, pill: COLORS.greenSoft, text: rgb(0.09, 0.42, 0.21) };
  if (value >= 70) return { track: COLORS.panelStrong, bar: COLORS.blue, pill: COLORS.panelStrong, text: COLORS.blue };
  if (value >= 40) return { track: COLORS.yellowSoft, bar: COLORS.yellow, pill: COLORS.yellowSoft, text: rgb(0.62, 0.39, 0.05) };
  return { track: COLORS.redSoft, bar: COLORS.red, pill: COLORS.redSoft, text: COLORS.red };
};

const drawProgressRow = (
  page: PDFPage,
  x: number,
  yTop: number,
  width: number,
  label: string,
  detail: string,
  value: number,
  fonts: FontSet,
) => {
  const tone = progressTone(value);
  page.drawText(sanitizePdfText(label), { x, y: yTop - 12, size: 8.5, font: fonts.bold, color: COLORS.ink });
  page.drawText(sanitizePdfText(detail), { x, y: yTop - 26, size: 7.1, font: fonts.body, color: COLORS.inkSoft });
  page.drawRectangle({ x: x + width - 48, y: yTop - 20, width: 48, height: 16, color: tone.pill });
  const pct = formatPercent(value);
  const pctWidth = fonts.bold.widthOfTextAtSize(pct, 7.8);
  page.drawText(pct, { x: x + width - 24 - pctWidth / 2, y: yTop - 14, size: 7.8, font: fonts.bold, color: tone.text });
  page.drawRectangle({ x, y: yTop - 40, width, height: 6, color: tone.track });
  page.drawRectangle({ x, y: yTop - 40, width: Math.max(24, width * Math.max(0, Math.min(1, value / 100))), height: 6, color: tone.bar });
};

const drawMiniStat = (
  page: PDFPage,
  x: number,
  yTop: number,
  width: number,
  label: string,
  value: string,
  note: string,
  fonts: FontSet,
  height = 72,
) => {
  page.drawRectangle({ x, y: yTop - height, width, height, color: COLORS.white, borderColor: COLORS.panelBorder, borderWidth: 1 });
  page.drawText(sanitizePdfText(label).toUpperCase(), {
    x: x + 12,
    y: yTop - 18,
    size: 7.1,
    font: fonts.bold,
    color: COLORS.inkSoft,
    characterSpacing: 1.1,
  });
  page.drawText(sanitizePdfText(value), {
    x: x + 12,
    y: yTop - 41,
    size: 16,
    font: fonts.bold,
    color: COLORS.ink,
  });
  drawTextBlock(page, note, x + 12, yTop - 55, width - 24, fonts.body, 6.8, COLORS.inkSoft, 2.1);
};

const drawTable = (
  page: PDFPage,
  x: number,
  yTop: number,
  width: number,
  columns: Array<{ label: string; width: number; align?: 'left' | 'right' }>,
  rows: string[][],
  fonts: FontSet,
  rowHeight = 28,
  bodySize = 8.3,
  headerSize = 7.8,
) => {
  page.drawRectangle({ x, y: yTop - TABLE_HEADER_HEIGHT, width, height: TABLE_HEADER_HEIGHT, color: COLORS.panel });
  let colX = x;
  columns.forEach(col => {
    const safeLabel = sanitizePdfText(col.label);
    const labelWidth = fonts.bold.widthOfTextAtSize(safeLabel, headerSize);
    const textX = col.align === 'right' ? colX + col.width - 10 - labelWidth : colX + 10;
    page.drawText(safeLabel, { x: textX, y: yTop - 17, size: headerSize, font: fonts.bold, color: COLORS.inkSoft, characterSpacing: col.align === 'right' ? 0.6 : 1.0 });
    colX += col.width;
  });
  page.drawLine({ start: { x, y: yTop - TABLE_HEADER_HEIGHT }, end: { x: x + width, y: yTop - TABLE_HEADER_HEIGHT }, thickness: 1, color: COLORS.line });

  rows.forEach((row, rowIndex) => {
    const top = yTop - TABLE_HEADER_HEIGHT - rowIndex * rowHeight;
    if (rowIndex > 0) {
      page.drawLine({ start: { x, y: top }, end: { x: x + width, y: top }, thickness: 1, color: COLORS.line });
    }
    let cellX = x;
    row.forEach((cell, idx) => {
      const safeCell = sanitizePdfText(cell);
      const align = columns[idx]?.align ?? 'left';
      if (align === 'right') {
        const textWidth = fonts.body.widthOfTextAtSize(safeCell, bodySize);
        page.drawText(safeCell, { x: cellX + columns[idx].width - 10 - textWidth, y: top - 18, size: bodySize, font: fonts.body, color: COLORS.ink });
      } else {
        page.drawText(safeCell, { x: cellX + 10, y: top - 18, size: bodySize, font: fonts.body, color: COLORS.ink });
      }
      cellX += columns[idx].width;
    });
  });
};

const drawFooter = (page: PDFPage, textLeft: string, textRight: string, fonts: FontSet, pageLabel: string) => {
  drawRule(page, PAGE.marginBottom + 10);
  page.drawText(sanitizePdfText(textLeft), { x: PAGE.marginX, y: FOOTER_Y, size: 8.5, font: fonts.body, color: COLORS.inkSoft });
  const right = sanitizePdfText(`${textRight} · ${pageLabel}`);
  const rightWidth = fonts.body.widthOfTextAtSize(right, 8.5);
  page.drawText(right, { x: PAGE.width - PAGE.marginX - rightWidth, y: FOOTER_Y, size: 8.5, font: fonts.body, color: COLORS.inkSoft });
};

type TitleOptions = {
  titleSize?: number;
  titleWidth?: number;
  maxTitleLines?: number;
};

const drawPageTitle = (
  page: PDFPage,
  title: string,
  subtitle: string,
  fonts: FontSet,
  options: TitleOptions = {},
) => {
  const topY = PAGE.height - PAGE.marginTop;
  page.drawText('MONIEZI TAX PREP PACKAGE', {
    x: PAGE.marginX,
    y: topY,
    size: 10.8,
    font: fonts.kicker,
    color: COLORS.blue,
    characterSpacing: 1.1,
  });

  const titleSize = options.titleSize ?? 26;
  const titleWidth = options.titleWidth ?? CONTENT_WIDTH;
  const titleLines = splitLines(title, fonts.bold, titleSize, titleWidth).slice(0, options.maxTitleLines ?? 2);
  const titleLineGap = Math.max(6, titleSize * 0.22);
  let titleY = topY - 56;
  titleLines.forEach(line => {
    page.drawText(sanitizePdfText(line), {
      x: PAGE.marginX,
      y: titleY,
      size: titleSize,
      font: fonts.bold,
      color: COLORS.ink,
    });
    titleY -= titleSize + titleLineGap;
  });

  const subtitleTop = titleY + 6;
  return drawTextBlock(page, subtitle, PAGE.marginX, subtitleTop, titleWidth, fonts.body, 8.9, COLORS.inkSoft, 3.5);
};

const drawInfoBoxRow = (
  page: PDFPage,
  yTop: number,
  blocks: Array<{ label: string; value: string; width: number; height?: number }>,
  fonts: FontSet,
  align: 'left' | 'right' = 'left',
  gap = 12,
) => {
  const totalWidth = blocks.reduce((sum, block) => sum + block.width, 0) + Math.max(0, blocks.length - 1) * gap;
  let x = align === 'right' ? PAGE.width - PAGE.marginX - totalWidth : PAGE.marginX;
  const maxHeight = Math.max(...blocks.map(block => block.height ?? 46));
  blocks.forEach(block => {
    drawInfoBox(page, x, yTop, block.width, block.label, block.value, fonts, block.height ?? 46);
    x += block.width + gap;
  });
  return yTop - maxHeight - 16;
};

const getPreparedByValue = (data: TaxSummaryPdfData) => {
  const owner = sanitizePdfText(data.ownerName);
  if (!owner || owner.toLowerCase() === 'owner') return 'Prepared privately';
  return owner;
};

export async function generateTaxSummaryPdfBytes(data: TaxSummaryPdfData): Promise<Uint8Array> {
  const pdfDoc = await PDFDocument.create();
  pdfDoc.registerFontkit(fontkit);

  let fonts: FontSet;
  try {
    const [reportRegularOtf, reportBoldOtf, appRegularOtf] = await Promise.all([
      getEmbeddedReportRegularOtf(),
      getEmbeddedReportBoldOtf(),
      getEmbeddedAppRegularOtf(),
    ]);
    fonts = {
      body: await pdfDoc.embedFont(reportRegularOtf, { subset: false }),
      bold: await pdfDoc.embedFont(reportBoldOtf, { subset: false }),
      kicker: await pdfDoc.embedFont(appRegularOtf, { subset: false }),
    };
  } catch (error) {
    console.warn('Tax Summary PDF custom fonts unavailable; falling back to standard PDF fonts.', error);
    fonts = {
      body: await pdfDoc.embedFont(StandardFonts.Helvetica),
      bold: await pdfDoc.embedFont(StandardFonts.HelveticaBold),
      kicker: await pdfDoc.embedFont(StandardFonts.HelveticaBold),
    };
  }

  const page1 = pdfDoc.addPage([PAGE.width, PAGE.height]);
  const page2 = pdfDoc.addPage([PAGE.width, PAGE.height]);
  const page3 = pdfDoc.addPage([PAGE.width, PAGE.height]);
  const page4 = pdfDoc.addPage([PAGE.width, PAGE.height]);

  const infoCardHeight = 86;
  const cardGap = 12;
  const cardWidth = (CONTENT_WIDTH - cardGap) / 2;

  const summaryRows = [
    { key: 'Gross business income', note: 'Recorded income transactions for the selected tax year.', value: formatCurrency(data.currencySymbol, data.totalIncome) },
    { key: 'Deductible business expenses', note: `${formatNumber(data.expenseItemsCount)} expense ${data.expenseItemsCount === 1 ? 'entry' : 'entries'} tracked before any external adjustments.`, value: formatCurrency(data.currencySymbol, data.totalExpenses) },
    { key: 'Net business profit', note: 'Income less recorded expenses for the selected period.', value: formatCurrency(data.currencySymbol, data.netProfit) },
    { key: 'Ledger transactions included', note: 'Total income and expense records represented in this package.', value: formatNumber(data.ledgerTransactions) },
    { key: 'Linked receipts attached', note: 'Receipt-backed expense records currently linked inside MONIEZI.', value: formatNumber(data.linkedReceipts) },
    { key: 'Expense categories used', note: 'Distinct deductible categories used during the selected tax year.', value: formatNumber(data.expenseCategoriesCount) },
  ];

  const progressRows = [
    { label: 'Receipt coverage', detail: `${formatNumber(data.linkedReceipts)} linked receipts across ${formatNumber(data.expenseItemsCount)} deductible expense items.`, value: data.receiptCoveragePct },
    { label: 'Expense review status', detail: `${formatNumber(data.reviewedExpenseCount)} reviewed · ${formatNumber(data.pendingReviewCount)} pending review.`, value: data.reviewCoveragePct },
    { label: 'Mileage log completeness', detail: `${formatNumber(data.completeMileageCount)} complete trip ${data.completeMileageCount === 1 ? 'entry' : 'entries'} recorded for ${formatNumber(data.totalMiles, 1)} business miles.`, value: data.mileageCompletionPct },
  ];

  const expenseRows = data.expenseRows.length
    ? data.expenseRows.map(row => [row.name, formatCurrency(data.currencySymbol, row.amount), formatPercent(row.sharePct), `${row.linked}/${row.count}`])
    : [['No deductible expenses were recorded for this tax year.', '', '', '']];

  const mileageRows = data.quarterlyMileage.length
    ? data.quarterlyMileage.map(row => [row.quarter, formatNumber(row.trips), formatNumber(row.miles, 1), formatCurrency(data.currencySymbol, row.deduction)])
    : [['Q1', '0', '0.0', formatCurrency(data.currencySymbol, 0)], ['Q2', '0', '0.0', formatCurrency(data.currencySymbol, 0)], ['Q3', '0', '0.0', formatCurrency(data.currencySymbol, 0)], ['Q4', '0', '0.0', formatCurrency(data.currencySymbol, 0)]];

  const attentionItems = data.attentionItems.length
    ? data.attentionItems
    : ['No major data gaps were detected in this tax-prep package.'];

  const preparedBy = getPreparedByValue(data);

  // PAGE 1
  let y = drawPageTitle(
    page1,
    `Executive Tax Snapshot ${sanitizePdfText(data.taxYear)}`,
    'A concise year-end package summarizing income, deductions, mileage, and supporting documentation from your MONIEZI records.',
    fonts,
    { titleSize: 24, maxTitleLines: 2 },
  );

  y = drawInfoBoxRow(
    page1,
    y - 8,
    [
      { label: 'Business', value: data.businessName, width: 162, height: 50 },
      { label: 'Tax year', value: data.taxYear, width: 162, height: 50 },
      { label: 'Reporting period', value: data.reportingPeriodLabel, width: 167, height: 50 },
    ],
    fonts,
  );

  const page1Cards = [
    { label: 'Gross business income', value: formatCurrency(data.currencySymbol, data.totalIncome), note: 'Recorded income transactions included in this export.' },
    { label: 'Deductible expenses', value: formatCurrency(data.currencySymbol, data.totalExpenses), note: `${formatNumber(data.expenseItemsCount)} tracked expense ${data.expenseItemsCount === 1 ? 'entry' : 'entries'} for the year.` },
    { label: 'Net business profit', value: formatCurrency(data.currencySymbol, data.netProfit), note: 'Income less recorded expenses before outside tax adjustments.' },
    { label: 'Mileage deduction', value: formatCurrency(data.currencySymbol, data.mileageDeduction), note: `${formatNumber(data.totalMiles, 1)} business miles at ${data.currencySymbol}${formatNumber(data.mileageRate, 2)} per mile.` },
  ];

  const cardsTop = y - 10;
  page1Cards.forEach((card, index) => {
    const col = index % 2;
    const row = Math.floor(index / 2);
    drawMetricCard(page1, PAGE.marginX + col * (cardWidth + cardGap), cardsTop - row * (infoCardHeight + cardGap), cardWidth, infoCardHeight, card.label, card.value, card.note, fonts);
  });

  const cardsBottom = cardsTop - (infoCardHeight * 2) - cardGap;
  const rowHeight = 36;
  const section1Height = SECTION_HEADER_HEIGHT + SECTION_INSET + summaryRows.length * rowHeight + SECTION_INSET;
  const section1Top = cardsBottom - 22;
  const section1BodyTop = drawSectionCard(
    page1,
    PAGE.marginX,
    section1Top,
    CONTENT_WIDTH,
    section1Height,
    '1',
    'Financial Snapshot & Package Scope',
    'Core figures, filing period details, and package metadata usually reviewed first before deeper tax preparation work begins.',
    fonts,
  );

  summaryRows.forEach((row, index) => {
    const rowTop = section1BodyTop - index * rowHeight;
    if (index > 0) {
      page1.drawLine({ start: { x: PAGE.marginX, y: rowTop }, end: { x: PAGE.marginX + CONTENT_WIDTH, y: rowTop }, thickness: 1, color: COLORS.line });
    }
    page1.drawText(sanitizePdfText(row.key), { x: PAGE.marginX + 14, y: rowTop - 16, size: 8.4, font: fonts.bold, color: COLORS.ink });
    page1.drawText(sanitizePdfText(row.note), { x: PAGE.marginX + 14, y: rowTop - 28, size: 7.1, font: fonts.body, color: COLORS.inkSoft });
    const valueWidth = fonts.bold.widthOfTextAtSize(sanitizePdfText(row.value), 9.3);
    page1.drawText(sanitizePdfText(row.value), { x: PAGE.marginX + CONTENT_WIDTH - 14 - valueWidth, y: rowTop - 17, size: 9.3, font: fonts.bold, color: COLORS.ink });
  });

  drawFooter(page1, 'MONIEZI Pro Finance — Generated privately from your local business records.', `${data.businessName} — Tax Year ${data.taxYear}`, fonts, 'Page 1 of 4');

  // PAGE 2
  y = drawPageTitle(
    page2,
    'Documentation Readiness',
    'A practical review of receipt coverage, expense review status, mileage completeness, and package readiness before filing.',
    fonts,
  );

  const section2Top = y - 14;
  const progressBlockHeight = progressRows.length * 50 + 16;
  const miniWidth = (CONTENT_WIDTH - 36 - 24) / 3;
  const miniStatHeight = 72;
  const section2Height = SECTION_HEADER_HEIGHT + SECTION_INSET + progressBlockHeight + miniStatHeight + 16 + SECTION_INSET;
  const section2BodyTop = drawSectionCard(
    page2,
    PAGE.marginX,
    section2Top,
    CONTENT_WIDTH,
    section2Height,
    '2',
    'Audit Readiness & Documentation Status',
    'A quick view of receipt coverage, review status, mileage completeness, and package readiness before filing.',
    fonts,
  );

  let progressY = section2BodyTop;
  progressRows.forEach((row, index) => {
    drawProgressRow(page2, PAGE.marginX + 18, progressY, CONTENT_WIDTH - 36, row.label, row.detail, row.value, fonts);
    progressY -= 50;
    if (index < progressRows.length - 1) progressY -= 10;
  });

  const page2MiniTop = progressY - 10;
  drawMiniStat(page2, PAGE.marginX + 18, page2MiniTop, miniWidth, 'Package coverage', formatNumber(data.ledgerTransactions), 'Ledger transactions included in this export package.', fonts, 72);
  drawMiniStat(page2, PAGE.marginX + 18 + miniWidth + 12, page2MiniTop, miniWidth, 'Items requiring attention', formatNumber(data.itemsRequiringAttention), 'Open items across receipts, review status, categories, and mileage.', fonts, 72);
  drawMiniStat(page2, PAGE.marginX + 18 + (miniWidth + 12) * 2, page2MiniTop, miniWidth, 'Prepared privately', '100%', 'Generated directly from MONIEZI local records for export and review.', fonts, 72);

  const page2NoteTop = section2Top - section2Height - 18;
  const page2NoteHeight = 88;
  page2.drawRectangle({ x: PAGE.marginX, y: page2NoteTop - page2NoteHeight, width: CONTENT_WIDTH, height: page2NoteHeight, color: COLORS.panel, borderColor: COLORS.panelBorder, borderWidth: 1 });
  page2.drawText('Readiness Takeaway', { x: PAGE.marginX + 18, y: page2NoteTop - 26, size: 12.4, font: fonts.bold, color: COLORS.ink });
  drawTextBlock(page2, `This page is designed to tell you, at a glance, how filing-ready your MONIEZI records look for tax year ${sanitizePdfText(data.taxYear)}. Low receipt or review coverage does not change your raw totals, but it does show where documentation or cleanup work may still be needed before filing.`, PAGE.marginX + 18, page2NoteTop - 40, CONTENT_WIDTH - 36, fonts.body, 9, COLORS.inkSoft, 4);

  drawFooter(page2, 'MONIEZI Pro Finance — Generated privately from your local business records.', `${data.businessName} — Tax Year ${data.taxYear}`, fonts, 'Page 2 of 4');

  // PAGE 3
  y = drawPageTitle(
    page3,
    'Deductible Expense Breakdown',
    'Top expense categories by dollar amount, category share, and receipt-backed count for the selected tax year.',
    fonts,
  );

  const page3HeroTop = y - 8;
  const page3HeroWidth = (CONTENT_WIDTH - 24) / 3;
  drawMiniStat(page3, PAGE.marginX, page3HeroTop, page3HeroWidth, 'Top expense category', data.topExpenseCategoryName || 'None', data.topExpenseCategoryName ? `${formatCurrency(data.currencySymbol, data.topExpenseCategoryAmount)} · ${formatPercent(data.topExpenseCategorySharePct)} of total expenses.` : 'No expense activity recorded for this period.', fonts, 72);
  drawMiniStat(page3, PAGE.marginX + page3HeroWidth + 12, page3HeroTop, page3HeroWidth, 'Expense items tracked', formatNumber(data.expenseItemsCount), 'Count of deductible expense entries included in the package.', fonts, 72);
  drawMiniStat(page3, PAGE.marginX + (page3HeroWidth + 12) * 2, page3HeroTop, page3HeroWidth, 'Linked receipts', formatNumber(data.linkedReceipts), 'Receipt-backed entries attached to deductible expenses.', fonts, 72);

  const expenseRowHeight = 26;
  const section3Top = page3HeroTop - 92;
  const section3Height = SECTION_HEADER_HEIGHT + SECTION_INSET + TABLE_HEADER_HEIGHT + expenseRows.length * expenseRowHeight + SECTION_INSET;
  const section3BodyTop = drawSectionCard(
    page3,
    PAGE.marginX,
    section3Top,
    CONTENT_WIDTH,
    section3Height,
    '3',
    'Deductible Expense Breakdown',
    'Top expense categories by dollar amount, category share, and receipt-backed count.',
    fonts,
  );
  drawTable(page3, PAGE.marginX, section3BodyTop, CONTENT_WIDTH, [
    { label: 'Expense category', width: 250 },
    { label: 'Amount', width: 116, align: 'right' },
    { label: 'Share', width: 72, align: 'right' },
    { label: 'Receipts', width: 77, align: 'right' },
  ], expenseRows, fonts, expenseRowHeight, 8.2, 7.8);

  drawFooter(page3, 'MONIEZI Pro Finance — Generated privately from your local business records.', `${data.businessName} — Tax Year ${data.taxYear}`, fonts, 'Page 3 of 4');

  // PAGE 4
  y = drawPageTitle(
    page4,
    'Mileage, Filing Checks & Handoff',
    'Quarter-level mileage review, open attention items, and final pre-filing guidance for this package.',
    fonts,
    { titleSize: 22, maxTitleLines: 2 },
  );

  y = drawInfoBoxRow(
    page4,
    y - 8,
    [{ label: 'Reporting period', value: data.reportingPeriodLabel, width: 196, height: 50 }],
    fonts,
    'right',
  );

  const splitGap = 14;
  const splitWidth = (CONTENT_WIDTH - splitGap) / 2;
  const mileageRowHeight = 26;
  const section4Height = SECTION_HEADER_HEIGHT + SECTION_INSET + TABLE_HEADER_HEIGHT + mileageRows.length * mileageRowHeight + SECTION_INSET;
  const attentionLineHeight = 13;
  const attentionBodyHeight = Math.max(88, attentionItems.reduce((sum, item, idx) => sum + textBlockHeight(item, fonts.body, 9, splitWidth - 46, 3.5) + (idx < attentionItems.length - 1 ? 10 : 0), 0));
  const section5Height = SECTION_HEADER_HEIGHT + SECTION_INSET + attentionBodyHeight + SECTION_INSET;
  const topSectionsHeight = Math.max(section4Height, section5Height);
  const topSectionsY = y - 10;

  const section4BodyTop = drawSectionCard(
    page4,
    PAGE.marginX,
    topSectionsY,
    splitWidth,
    section4Height,
    '4',
    'Mileage Log Summary',
    'Quarter-by-quarter view of recorded trips, miles, and estimated deduction.',
    fonts,
  );
  drawTable(page4, PAGE.marginX, section4BodyTop, splitWidth, [
    { label: 'QTR', width: 58 },
    { label: 'Trips', width: 52, align: 'right' },
    { label: 'Miles', width: 76, align: 'right' },
    { label: 'Deduction', width: splitWidth - 58 - 52 - 76, align: 'right' },
  ], mileageRows, fonts, mileageRowHeight, 8.2, 7.6);

  const section5BodyTop = drawSectionCard(
    page4,
    PAGE.marginX + splitWidth + splitGap,
    topSectionsY,
    splitWidth,
    section5Height,
    '5',
    'Attention Items Before Filing',
    'Items that should be reviewed or completed before handing records to your tax preparer.',
    fonts,
  );
  let bulletY = section5BodyTop;
  attentionItems.forEach((item, index) => {
    const lines = splitLines(item, fonts.body, 9, splitWidth - 46);
    page4.drawCircle({ x: PAGE.marginX + splitWidth + splitGap + 16, y: bulletY + 3, size: 2.5, color: COLORS.blue });
    lines.forEach((line, lineIndex) => {
      page4.drawText(sanitizePdfText(line), {
        x: PAGE.marginX + splitWidth + splitGap + 26,
        y: bulletY - lineIndex * attentionLineHeight,
        size: 9,
        font: fonts.body,
        color: COLORS.ink,
      });
    });
    bulletY -= lines.length * attentionLineHeight + (index < attentionItems.length - 1 ? 10 : 0);
  });

  const noteText = `MONIEZI prepared this package from recorded ledger entries, linked receipts, and mileage logs for tax year ${sanitizePdfText(data.taxYear)}. Use it as a filing-ready summary of what was earned, what was spent, how well expenses are documented, and which cleanup items remain. Final tax treatment, classification decisions, and any required adjustments should still be reviewed with your tax professional.`;
  const noteHeight = 34 + textBlockHeight(noteText, fonts.body, 9.1, CONTENT_WIDTH - 36, 4.1) + 18;
  const noteTop = topSectionsY - topSectionsHeight - 18;
  page4.drawRectangle({ x: PAGE.marginX, y: noteTop - noteHeight, width: CONTENT_WIDTH, height: noteHeight, color: COLORS.panel, borderColor: COLORS.panelBorder, borderWidth: 1 });
  page4.drawText('Pre-Filing Note', { x: PAGE.marginX + 18, y: noteTop - 26, size: 12.2, font: fonts.bold, color: COLORS.ink });
  drawTextBlock(page4, noteText, PAGE.marginX + 18, noteTop - 40, CONTENT_WIDTH - 36, fonts.body, 9.1, COLORS.inkSoft, 4.1);

  const finalStatsTop = noteTop - noteHeight - 18;
  drawMiniStat(page4, PAGE.marginX, finalStatsTop, (CONTENT_WIDTH - 12) / 2, 'Prepared by', preparedBy, 'Name shown for package reference only. Data remains stored locally.', fonts, 64);
  drawMiniStat(page4, PAGE.marginX + (CONTENT_WIDTH - 12) / 2 + 12, finalStatsTop, (CONTENT_WIDTH - 12) / 2, 'Generated', data.generatedAtLabel, 'Export timestamp for this tax package.', fonts, 64);

  drawFooter(page4, 'MONIEZI Pro Finance — Generated privately from your local business records.', `${data.businessName} — Tax Year ${data.taxYear}`, fonts, 'Page 4 of 4');

  pdfDoc.setTitle(sanitizePdfText(`MONIEZI Tax Prep Package ${data.taxYear}`));
  pdfDoc.setAuthor('MONIEZI');
  pdfDoc.setCreator('MONIEZI Pro Finance');
  pdfDoc.setProducer('MONIEZI Pro Finance');
  pdfDoc.setSubject(sanitizePdfText(`Tax Prep Package ${data.taxYear}`));

  return pdfDoc.save();
}
