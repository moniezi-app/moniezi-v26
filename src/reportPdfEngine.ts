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

const PAGE = {
  width: 595.28,
  height: 841.89,
  marginX: 40,
  top: 46,
  bottom: 38,
};

const CONTENT_WIDTH = PAGE.width - PAGE.marginX * 2;
const SECTION_HEADER_HEIGHT = 64;
const SECTION_TOP_PAD = 18;
const SECTION_BOTTOM_PAD = 18;
const TABLE_HEADER_HEIGHT = 28;

const COLORS = {
  ink: rgb(0.07, 0.11, 0.2),
  inkSoft: rgb(0.38, 0.45, 0.56),
  line: rgb(0.86, 0.89, 0.94),
  page: rgb(1, 1, 1),
  panel: rgb(0.97, 0.98, 1),
  panelStrong: rgb(0.94, 0.96, 1),
  panelBorder: rgb(0.84, 0.89, 0.95),
  blue: rgb(0.16, 0.38, 0.89),
  blueSoft: rgb(0.93, 0.96, 1),
  green: rgb(0.15, 0.67, 0.34),
  greenSoft: rgb(0.9, 0.97, 0.92),
  red: rgb(0.86, 0.26, 0.26),
  redSoft: rgb(1, 0.93, 0.93),
  yellow: rgb(0.85, 0.59, 0.11),
  yellowSoft: rgb(1, 0.96, 0.86),
};

const sanitizePdfText = (value: unknown) => String(value ?? '')
  .normalize('NFKD')
  .replace(/[\u2010-\u2015]/g, '-')
  .replace(/[\u2022\u00B7]/g, '-')
  .replace(/[\u2018\u2019]/g, "'")
  .replace(/[\u201C\u201D]/g, '"')
  .replace(/\u2026/g, '...')
  .replace(/\u00A0/g, ' ')
  .replace(/â€[\x90-\xBF]?/g, '-')
  .replace(/Â·/g, '-')
  .replace(/[^\n\x20-\x7E]/g, ' ')
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

const measureBlockHeight = (text: string, font: PDFFont, size: number, width: number, lineGap = 4) => {
  const lines = splitLines(text, font, size, width);
  if (!lines.length) return 0;
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

const sectionHeight = (bodyHeight: number) => SECTION_HEADER_HEIGHT + SECTION_TOP_PAD + bodyHeight + SECTION_BOTTOM_PAD;

const drawPageHeader = (
  page: PDFPage,
  title: string,
  subtitle: string,
  bodyFont: PDFFont,
  boldFont: PDFFont,
  kickerFont: PDFFont,
) => {
  const topY = PAGE.height - PAGE.top;
  page.drawText('MONIEZI TAX PREP PACKAGE', {
    x: PAGE.marginX,
    y: topY,
    size: 10.8,
    font: kickerFont,
    color: COLORS.blue,
    characterSpacing: 1.2,
  });
  page.drawText(sanitizePdfText(title), {
    x: PAGE.marginX,
    y: topY - 54,
    size: 24,
    font: boldFont,
    color: COLORS.ink,
  });
  const afterSubtitle = drawTextBlock(page, subtitle, PAGE.marginX, topY - 74, 340, bodyFont, 8.8, COLORS.inkSoft, 3.5);
  return afterSubtitle - 8;
};

const drawMetaBlock = (
  page: PDFPage,
  x: number,
  yTop: number,
  width: number,
  label: string,
  value: string,
  bodyFont: PDFFont,
  boldFont: PDFFont,
  height = 42,
) => {
  page.drawRectangle({
    x,
    y: yTop - height,
    width,
    height,
    borderWidth: 1,
    borderColor: COLORS.panelBorder,
    color: COLORS.blueSoft,
  });
  page.drawText(sanitizePdfText(label).toUpperCase(), {
    x: x + 12,
    y: yTop - 16,
    size: 7.1,
    font: boldFont,
    color: COLORS.inkSoft,
    characterSpacing: 1.3,
  });
  const valueSize = 9.1;
  const lines = splitLines(value, boldFont, valueSize, width - 24);
  let cursor = yTop - 30;
  lines.slice(0, 2).forEach(line => {
    page.drawText(sanitizePdfText(line), { x: x + 12, y: cursor, size: valueSize, font: boldFont, color: COLORS.ink });
    cursor -= 11;
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
  bodyFont: PDFFont,
  boldFont: PDFFont,
) => {
  page.drawRectangle({ x, y: yTop - height, width, height, borderWidth: 1, borderColor: COLORS.panelBorder, color: COLORS.page });
  page.drawText(sanitizePdfText(label).toUpperCase(), {
    x: x + 14,
    y: yTop - 18,
    size: 7.2,
    font: boldFont,
    color: COLORS.inkSoft,
    characterSpacing: 1.2,
  });
  page.drawText(sanitizePdfText(value), {
    x: x + 14,
    y: yTop - 42,
    size: 15.5,
    font: boldFont,
    color: COLORS.ink,
  });
  drawTextBlock(page, note, x + 14, yTop - 56, width - 28, bodyFont, 6.9, COLORS.inkSoft, 2.1);
};

const drawSectionShell = (
  page: PDFPage,
  x: number,
  yTop: number,
  width: number,
  height: number,
  sectionNo: string,
  title: string,
  subtitle: string,
  bodyFont: PDFFont,
  boldFont: PDFFont,
) => {
  page.drawRectangle({ x, y: yTop - height, width, height, borderWidth: 1, borderColor: COLORS.panelBorder, color: COLORS.page });
  page.drawRectangle({ x, y: yTop - SECTION_HEADER_HEIGHT, width, height: SECTION_HEADER_HEIGHT, color: COLORS.panel });
  page.drawText(`SECTION ${sanitizePdfText(sectionNo)}`, {
    x: x + 18,
    y: yTop - 18,
    size: 7.8,
    font: boldFont,
    color: COLORS.blue,
    characterSpacing: 1.3,
  });
  page.drawText(sanitizePdfText(title), {
    x: x + 18,
    y: yTop - 40,
    size: 13.8,
    font: boldFont,
    color: COLORS.ink,
  });
  drawTextBlock(page, subtitle, x + 18, yTop - 50, width - 36, bodyFont, 7.2, COLORS.inkSoft, 2.2);
  page.drawLine({ start: { x, y: yTop - SECTION_HEADER_HEIGHT }, end: { x: x + width, y: yTop - SECTION_HEADER_HEIGHT }, thickness: 1, color: COLORS.line });
  return yTop - SECTION_HEADER_HEIGHT - SECTION_TOP_PAD;
};

interface DetailRow {
  key: string;
  note: string;
  value: string;
  emphasize?: boolean;
}

const measureDetailRowsHeight = (rows: DetailRow[], bodyFont: PDFFont, valueFont: PDFFont, width: number) => {
  return rows.reduce((sum, row) => {
    const noteWidth = width - 190;
    const noteHeight = measureBlockHeight(row.note, bodyFont, 7.3, noteWidth, 2.2);
    const valueLines = splitLines(row.value, valueFont, 9.6, 126);
    const valueHeight = valueLines.length * 9.6 + Math.max(0, valueLines.length - 1) * 2;
    const rowHeight = Math.max(44, 22 + Math.max(noteHeight + 10, valueHeight + 10));
    return sum + rowHeight;
  }, 0);
};

const drawDetailRows = (
  page: PDFPage,
  x: number,
  yTop: number,
  width: number,
  rows: DetailRow[],
  bodyFont: PDFFont,
  boldFont: PDFFont,
) => {
  let cursor = yTop;
  rows.forEach((row, idx) => {
    const noteWidth = width - 190;
    const noteHeight = measureBlockHeight(row.note, bodyFont, 7.3, noteWidth, 2.2);
    const valueLines = splitLines(row.value, boldFont, 9.6, 126);
    const valueHeight = valueLines.length * 9.6 + Math.max(0, valueLines.length - 1) * 2;
    const rowHeight = Math.max(44, 22 + Math.max(noteHeight + 10, valueHeight + 10));

    if (row.emphasize) {
      page.drawRectangle({ x, y: cursor - rowHeight, width, height: rowHeight, color: COLORS.blueSoft });
    }
    if (idx > 0) {
      page.drawLine({ start: { x, y: cursor }, end: { x: x + width, y: cursor }, thickness: 1, color: COLORS.line });
    }

    page.drawText(sanitizePdfText(row.key), {
      x: x + 14,
      y: cursor - 16,
      size: 8.8,
      font: boldFont,
      color: COLORS.ink,
    });
    drawTextBlock(page, row.note, x + 14, cursor - 24, noteWidth, bodyFont, 7.3, COLORS.inkSoft, 2.2);

    let valueY = cursor - 18;
    valueLines.forEach(line => {
      const safe = sanitizePdfText(line);
      const lineWidth = boldFont.widthOfTextAtSize(safe, 9.6);
      page.drawText(safe, { x: x + width - 14 - lineWidth, y: valueY, size: 9.6, font: boldFont, color: COLORS.ink });
      valueY -= 11.2;
    });

    cursor -= rowHeight;
  });
};

const progressTone = (value: number) => {
  if (value >= 90) return { track: COLORS.greenSoft, bar: COLORS.green, pill: COLORS.greenSoft, text: rgb(0.11, 0.42, 0.2) };
  if (value >= 70) return { track: COLORS.blueSoft, bar: COLORS.blue, pill: COLORS.blueSoft, text: COLORS.blue };
  if (value >= 40) return { track: COLORS.yellowSoft, bar: COLORS.yellow, pill: COLORS.yellowSoft, text: rgb(0.55, 0.36, 0.05) };
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
  bodyFont: PDFFont,
  boldFont: PDFFont,
) => {
  const tone = progressTone(value);
  page.drawText(sanitizePdfText(label), { x, y: yTop - 12, size: 8.4, font: boldFont, color: COLORS.ink });
  drawTextBlock(page, detail, x, yTop - 18, width - 64, bodyFont, 7.2, COLORS.inkSoft, 2.1);
  page.drawRectangle({ x: x + width - 50, y: yTop - 20, width: 50, height: 18, color: tone.pill });
  const pct = formatPercent(value);
  const pctWidth = boldFont.widthOfTextAtSize(pct, 8);
  page.drawText(pct, { x: x + width - 25 - pctWidth / 2, y: yTop - 14, size: 8, font: boldFont, color: tone.text });
  page.drawRectangle({ x, y: yTop - 40, width, height: 7, color: tone.track });
  page.drawRectangle({ x, y: yTop - 40, width: Math.max(20, width * Math.max(0, Math.min(1, value / 100))), height: 7, color: tone.bar });
};

const drawMiniStat = (
  page: PDFPage,
  x: number,
  yTop: number,
  width: number,
  height: number,
  label: string,
  value: string,
  note: string,
  bodyFont: PDFFont,
  boldFont: PDFFont,
) => {
  page.drawRectangle({ x, y: yTop - height, width, height, borderWidth: 1, borderColor: COLORS.panelBorder, color: COLORS.page });
  page.drawText(sanitizePdfText(label).toUpperCase(), { x: x + 12, y: yTop - 16, size: 7.2, font: boldFont, color: COLORS.inkSoft, characterSpacing: 0.9 });
  page.drawText(sanitizePdfText(value), { x: x + 12, y: yTop - 38, size: 15, font: boldFont, color: COLORS.ink });
  drawTextBlock(page, note, x + 12, yTop - 50, width - 24, bodyFont, 6.7, COLORS.inkSoft, 2);
};

const drawTable = (
  page: PDFPage,
  x: number,
  yTop: number,
  width: number,
  columns: Array<{ label: string; width: number; align?: 'left' | 'right' }>,
  rows: string[][],
  bodyFont: PDFFont,
  boldFont: PDFFont,
  rowHeight = 28,
  bodyFontSize = 8.4,
  headerFontSize = 7.9,
) => {
  page.drawRectangle({ x, y: yTop - TABLE_HEADER_HEIGHT, width, height: TABLE_HEADER_HEIGHT, color: COLORS.panel });
  let cellX = x;
  columns.forEach(col => {
    const safe = sanitizePdfText(col.label);
    const labelWidth = boldFont.widthOfTextAtSize(safe, headerFontSize);
    const textX = col.align === 'right' ? cellX + col.width - 12 - labelWidth : cellX + 12;
    page.drawText(safe, { x: textX, y: yTop - 18, size: headerFontSize, font: boldFont, color: COLORS.inkSoft, characterSpacing: 0.8 });
    cellX += col.width;
  });
  page.drawLine({ start: { x, y: yTop - TABLE_HEADER_HEIGHT }, end: { x: x + width, y: yTop - TABLE_HEADER_HEIGHT }, thickness: 1, color: COLORS.line });

  rows.forEach((row, rowIndex) => {
    const rowTop = yTop - TABLE_HEADER_HEIGHT - rowIndex * rowHeight;
    if (rowIndex > 0) {
      page.drawLine({ start: { x, y: rowTop }, end: { x: x + width, y: rowTop }, thickness: 1, color: COLORS.line });
    }
    let currentX = x;
    row.forEach((cell, idx) => {
      const safe = sanitizePdfText(cell);
      const col = columns[idx];
      if (col.align === 'right') {
        const textWidth = bodyFont.widthOfTextAtSize(safe, bodyFontSize);
        page.drawText(safe, { x: currentX + col.width - 12 - textWidth, y: rowTop - 17, size: bodyFontSize, font: bodyFont, color: COLORS.ink });
      } else {
        page.drawText(safe, { x: currentX + 12, y: rowTop - 17, size: bodyFontSize, font: bodyFont, color: COLORS.ink });
      }
      currentX += col.width;
    });
  });
};

const drawBulletList = (
  page: PDFPage,
  x: number,
  yTop: number,
  width: number,
  items: string[],
  bodyFont: PDFFont,
) => {
  let cursor = yTop;
  items.forEach((item, idx) => {
    const lines = splitLines(item, bodyFont, 8.8, width - 24);
    page.drawCircle({ x: x + 4, y: cursor - 4, size: 2.2, color: idx === 0 ? COLORS.blue : COLORS.inkSoft });
    lines.forEach((line, lineIndex) => {
      page.drawText(sanitizePdfText(line), { x: x + 14, y: cursor - lineIndex * 12, size: 8.8, font: bodyFont, color: COLORS.ink });
    });
    cursor -= lines.length * 12 + 8;
  });
};

const drawFooter = (page: PDFPage, pageNo: number, totalPages: number, businessName: string, taxYear: string, bodyFont: PDFFont, boldFont: PDFFont) => {
  const lineY = PAGE.bottom;
  page.drawLine({ start: { x: PAGE.marginX, y: lineY }, end: { x: PAGE.width - PAGE.marginX, y: lineY }, thickness: 1, color: COLORS.line });
  const left = 'MONIEZI Pro Finance - Generated privately from your local business records.';
  const right = `${sanitizePdfText(businessName)} - Tax Year ${sanitizePdfText(taxYear)} - Page ${pageNo} of ${totalPages}`;
  page.drawText(sanitizePdfText(left), { x: PAGE.marginX, y: lineY - 12, size: 8.1, font: bodyFont, color: COLORS.inkSoft });
  const rightWidth = bodyFont.widthOfTextAtSize(sanitizePdfText(right), 8.1);
  page.drawText(sanitizePdfText(right), { x: PAGE.width - PAGE.marginX - rightWidth, y: lineY - 12, size: 8.1, font: bodyFont, color: COLORS.inkSoft });
};

export async function generateTaxSummaryPdfBytes(data: TaxSummaryPdfData): Promise<Uint8Array> {
  const pdfDoc = await PDFDocument.create();
  pdfDoc.registerFontkit(fontkit);

  let bodyFont: PDFFont;
  let boldFont: PDFFont;
  let kickerFont: PDFFont;

  try {
    const [reportRegularOtf, reportBoldOtf, appRegularOtf] = await Promise.all([
      getEmbeddedReportRegularOtf(),
      getEmbeddedReportBoldOtf(),
      getEmbeddedAppRegularOtf(),
    ]);
    bodyFont = await pdfDoc.embedFont(reportRegularOtf, { subset: false });
    boldFont = await pdfDoc.embedFont(reportBoldOtf, { subset: false });
    kickerFont = await pdfDoc.embedFont(appRegularOtf, { subset: false });
  } catch (error) {
    console.warn('Tax Summary PDF custom fonts unavailable; falling back to standard PDF fonts.', error);
    bodyFont = await pdfDoc.embedFont(StandardFonts.Helvetica);
    boldFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
    kickerFont = boldFont;
  }

  const page1 = pdfDoc.addPage([PAGE.width, PAGE.height]);
  const page2 = pdfDoc.addPage([PAGE.width, PAGE.height]);
  const page3 = pdfDoc.addPage([PAGE.width, PAGE.height]);
  const pages = [page1, page2, page3];

  const summaryCards = [
    { label: 'Gross Business Income', value: formatCurrency(data.currencySymbol, data.totalIncome), note: `Recorded income transactions included for tax year ${data.taxYear}.` },
    { label: 'Deductible Expenses', value: formatCurrency(data.currencySymbol, data.totalExpenses), note: `${formatNumber(data.expenseItemsCount)} expense entries tracked before outside adjustments.` },
    { label: 'Net Business Profit', value: formatCurrency(data.currencySymbol, data.netProfit), note: 'Income less recorded deductible expenses for the selected tax year.' },
    { label: 'Mileage Deduction', value: formatCurrency(data.currencySymbol, data.mileageDeduction), note: `${formatNumber(data.totalMiles, 1)} business miles at ${data.currencySymbol}${formatNumber(data.mileageRate, 2)} per mile.` },
  ];

  const section1Rows: DetailRow[] = [
    { key: 'Gross Business Income', note: 'Total recorded income transactions captured inside MONIEZI for the selected tax year.', value: formatCurrency(data.currencySymbol, data.totalIncome) },
    { key: 'Deductible Business Expenses', note: 'Total expense entries included in this package before any accountant-side adjustments.', value: formatCurrency(data.currencySymbol, data.totalExpenses) },
    { key: 'Net Business Profit', note: 'Recorded income less recorded deductible expenses for this reporting year.', value: formatCurrency(data.currencySymbol, data.netProfit), emphasize: true },
    { key: 'Mileage Logged', note: `${formatNumber(data.completeMileageCount)} complete trip ${data.completeMileageCount === 1 ? 'entry' : 'entries'} captured in the mileage log.`, value: `${formatNumber(data.totalMiles, 1)} mi` },
    { key: 'Mileage Rate Used', note: 'Configured standard mileage rate applied to the exported tax package.', value: `${data.currencySymbol}${formatNumber(data.mileageRate, 2)} per mi` },
    { key: 'Estimated Mileage Deduction', note: 'Computed directly from recorded business mileage using the configured rate.', value: formatCurrency(data.currencySymbol, data.mileageDeduction) },
    { key: 'Ledger Transactions Included', note: 'Income and expense records packaged into this year-end export.', value: formatNumber(data.ledgerTransactions) },
    { key: 'Reporting Period', note: 'Earliest through latest dated record included in this package.', value: data.reportingPeriodLabel },
  ];

  const progressRows = [
    { label: 'Receipt Coverage', detail: `${formatNumber(data.linkedReceipts)} linked receipts across ${formatNumber(data.expenseItemsCount)} deductible expense items.`, value: data.receiptCoveragePct },
    { label: 'Expense Review Status', detail: `${formatNumber(data.reviewedExpenseCount)} reviewed, ${formatNumber(data.pendingReviewCount)} pending review.`, value: data.reviewCoveragePct },
    { label: 'Mileage Log Completeness', detail: `${formatNumber(data.completeMileageCount)} complete trip entries recorded for ${formatNumber(data.totalMiles, 1)} business miles.`, value: data.mileageCompletionPct },
  ];

  const expenseRows = data.expenseRows.length
    ? data.expenseRows.map(row => [row.name, formatCurrency(data.currencySymbol, row.amount), formatPercent(row.sharePct), `${row.linked}/${row.count}`])
    : [['No deductible expenses were recorded for this tax year.', '', '', '']];

  const mileageRows = data.quarterlyMileage.map(row => [row.quarter.replace('Quarter', 'QTR').replace(/^Q(\d)$/, 'Q$1'), formatNumber(row.trips), `${formatNumber(row.miles, 1)} mi`, formatCurrency(data.currencySymbol, row.deduction)]);

  const attentionItems = data.attentionItems.length
    ? data.attentionItems
    : ['No major data gaps were detected in this tax-prep package.'];

  const closingNote = 'This package was prepared from your recorded ledger entries, linked receipt attachments, and mileage logs for the selected tax year. It is designed to show what was earned, what was spent, how well expenses are documented, and which items still need attention before filing. Final tax treatment, classification decisions, and any adjustments should still be reviewed with your tax professional.';

  // Page 1
  let y = drawPageHeader(
    page1,
    `Tax Prep Package Summary ${sanitizePdfText(data.taxYear)}`,
    'Year-end financial totals, package scope, and core tax-ready figures prepared from your MONIEZI business records.',
    bodyFont,
    boldFont,
    kickerFont,
  );
  drawMetaBlock(page1, PAGE.marginX + CONTENT_WIDTH - 154, PAGE.height - PAGE.top + 6, 154, 'Business', data.businessName || 'Business', bodyFont, boldFont, 48);
  drawMetaBlock(page1, PAGE.marginX + CONTENT_WIDTH - 154, PAGE.height - PAGE.top - 52, 154, 'Tax Year', data.taxYear, bodyFont, boldFont, 42);
  drawMetaBlock(page1, PAGE.marginX + CONTENT_WIDTH - 154, PAGE.height - PAGE.top - 102, 154, 'Reporting Period', data.reportingPeriodLabel, bodyFont, boldFont, 50);
  drawMetaBlock(page1, PAGE.marginX + CONTENT_WIDTH - 154, PAGE.height - PAGE.top - 160, 154, 'Generated', data.generatedAtLabel, bodyFont, boldFont, 42);

  const cardGap = 12;
  const cardWidth = (CONTENT_WIDTH - cardGap * 3) / 4;
  const cardHeight = 80;
  summaryCards.forEach((card, idx) => {
    drawMetricCard(page1, PAGE.marginX + idx * (cardWidth + cardGap), y - 4, cardWidth, cardHeight, card.label, card.value, card.note, bodyFont, boldFont);
  });

  const section1Height = sectionHeight(measureDetailRowsHeight(section1Rows, bodyFont, boldFont, CONTENT_WIDTH));
  const section1Top = y - cardHeight - 18;
  const section1BodyTop = drawSectionShell(
    page1,
    PAGE.marginX,
    section1Top,
    CONTENT_WIDTH,
    section1Height,
    '1',
    'Tax-Ready Financial Snapshot',
    'Core year-end amounts and package details your accountant typically needs first.',
    bodyFont,
    boldFont,
  );
  drawDetailRows(page1, PAGE.marginX, section1BodyTop, CONTENT_WIDTH, section1Rows, bodyFont, boldFont);

  // Page 2
  y = drawPageHeader(
    page2,
    'Documentation & Deduction Detail',
    'Readiness indicators, package coverage, and the highest-impact deductible categories for the selected tax year.',
    bodyFont,
    boldFont,
    kickerFont,
  );
  const metaWidth = 116;
  const metaGap = 10;
  const metaStartX = PAGE.marginX + CONTENT_WIDTH - (metaWidth * 3 + metaGap * 2);
  drawMetaBlock(page2, metaStartX, PAGE.height - PAGE.top + 6, metaWidth, 'Tax Year', data.taxYear, bodyFont, boldFont, 40);
  drawMetaBlock(page2, metaStartX + metaWidth + metaGap, PAGE.height - PAGE.top + 6, metaWidth, 'Linked Receipts', formatNumber(data.linkedReceipts), bodyFont, boldFont, 40);
  drawMetaBlock(page2, metaStartX + (metaWidth + metaGap) * 2, PAGE.height - PAGE.top + 6, metaWidth, 'Expense Categories', formatNumber(data.expenseCategoriesCount), bodyFont, boldFont, 40);

  const progressContentHeight = progressRows.length * 50 + (progressRows.length - 1) * 12 + 14 + 82;
  const section2Height = sectionHeight(progressContentHeight);
  const section2Top = y - 8;
  const section2BodyTop = drawSectionShell(
    page2,
    PAGE.marginX,
    section2Top,
    CONTENT_WIDTH,
    section2Height,
    '2',
    'Audit Readiness & Documentation Status',
    'A quick view of receipt coverage, review status, mileage completeness, and package readiness before filing.',
    bodyFont,
    boldFont,
  );

  let progressY = section2BodyTop;
  progressRows.forEach((row, idx) => {
    drawProgressRow(page2, PAGE.marginX + 18, progressY, CONTENT_WIDTH - 36, row.label, row.detail, row.value, bodyFont, boldFont);
    progressY -= 50;
    if (idx < progressRows.length - 1) progressY -= 12;
  });

  const miniTop = progressY - 12;
  const miniHeight = 74;
  const miniGap = 12;
  const miniWidth = (CONTENT_WIDTH - 36 - miniGap * 2) / 3;
  drawMiniStat(page2, PAGE.marginX + 18, miniTop, miniWidth, miniHeight, 'Package Coverage', formatNumber(data.ledgerTransactions), 'Total ledger transactions included in this tax-prep package export.', bodyFont, boldFont);
  drawMiniStat(page2, PAGE.marginX + 18 + miniWidth + miniGap, miniTop, miniWidth, miniHeight, 'Items Requiring Attention', formatNumber(data.itemsRequiringAttention), 'Headline open items across receipts, review status, categorization, and mileage completeness.', bodyFont, boldFont);
  drawMiniStat(page2, PAGE.marginX + 18 + (miniWidth + miniGap) * 2, miniTop, miniWidth, miniHeight, 'Top Expense Category', data.topExpenseCategoryName || 'None recorded', data.topExpenseCategoryName ? `${formatCurrency(data.currencySymbol, data.topExpenseCategoryAmount)} or ${formatPercent(data.topExpenseCategorySharePct)} of deductible expenses.` : 'No expense category rose to the top in this period.', bodyFont, boldFont);

  const section3BodyHeight = TABLE_HEADER_HEIGHT + expenseRows.length * 28;
  const section3Height = sectionHeight(section3BodyHeight);
  const section3Top = section2Top - section2Height - 16;
  const section3BodyTop = drawSectionShell(
    page2,
    PAGE.marginX,
    section3Top,
    CONTENT_WIDTH,
    section3Height,
    '3',
    'Deductible Expense Breakdown',
    'Top expense categories by dollar amount, category share, and receipt-backed count.',
    bodyFont,
    boldFont,
  );
  drawTable(
    page2,
    PAGE.marginX,
    section3BodyTop,
    CONTENT_WIDTH,
    [
      { label: 'Expense Category', width: 250 },
      { label: 'Amount', width: 120, align: 'right' },
      { label: 'Share', width: 70, align: 'right' },
      { label: 'Receipts', width: 75, align: 'right' },
    ],
    expenseRows,
    bodyFont,
    boldFont,
    28,
    8.3,
    7.8,
  );

  // Page 3
  y = drawPageHeader(
    page3,
    'Mileage, Filing Checks & Handoff',
    'Quarter-level mileage review, open attention items, and final pre-filing guidance for this package.',
    bodyFont,
    boldFont,
    kickerFont,
  );
  drawMetaBlock(page3, PAGE.marginX + CONTENT_WIDTH - 180, PAGE.height - PAGE.top + 6, 180, 'Reporting Period', data.reportingPeriodLabel, bodyFont, boldFont, 48);

  const splitGap = 16;
  const splitWidth = (CONTENT_WIDTH - splitGap) / 2;
  const mileageBodyHeight = data.hasMileageRows ? TABLE_HEADER_HEIGHT + mileageRows.length * 28 : 72;
  const section4Height = sectionHeight(mileageBodyHeight);
  const attentionBodyHeight = Math.max(64, attentionItems.reduce((sum, item) => sum + splitLines(item, bodyFont, 8.8, splitWidth - 42).length * 12 + 8, 0));
  const section5Height = sectionHeight(attentionBodyHeight);
  const sectionTop = y - 8;

  const section4BodyTop = drawSectionShell(
    page3,
    PAGE.marginX,
    sectionTop,
    splitWidth,
    section4Height,
    '4',
    'Mileage Log Summary',
    'Quarter-by-quarter view of recorded trips, miles, and estimated deduction.',
    bodyFont,
    boldFont,
  );
  if (data.hasMileageRows) {
    drawTable(
      page3,
      PAGE.marginX,
      section4BodyTop,
      splitWidth,
      [
        { label: 'QTR', width: 54 },
        { label: 'Trips', width: 54, align: 'right' },
        { label: 'Miles', width: 92, align: 'right' },
        { label: 'Deduction', width: splitWidth - 54 - 54 - 92, align: 'right' },
      ],
      mileageRows,
      bodyFont,
      boldFont,
      28,
      8.2,
      7.6,
    );
  } else {
    page3.drawText('No mileage trips were recorded for this tax year.', { x: PAGE.marginX + 18, y: section4BodyTop - 18, size: 9.2, font: bodyFont, color: COLORS.inkSoft });
    page3.drawText('The package still includes your configured mileage rate for reference.', { x: PAGE.marginX + 18, y: section4BodyTop - 34, size: 8.2, font: bodyFont, color: COLORS.inkSoft });
  }

  const section5BodyTop = drawSectionShell(
    page3,
    PAGE.marginX + splitWidth + splitGap,
    sectionTop,
    splitWidth,
    section5Height,
    '5',
    'Attention Items Before Filing',
    'Items that should be reviewed or completed before handing records to your tax preparer.',
    bodyFont,
    boldFont,
  );
  drawBulletList(page3, PAGE.marginX + splitWidth + splitGap + 18, section5BodyTop - 2, splitWidth - 36, attentionItems, bodyFont);

  const noteTextHeight = measureBlockHeight(closingNote, bodyFont, 8.8, CONTENT_WIDTH - 36, 3.2);
  const section6BodyHeight = noteTextHeight + 8;
  const section6Height = sectionHeight(section6BodyHeight);
  const section6Top = sectionTop - Math.max(section4Height, section5Height) - 16;
  const section6BodyTop = drawSectionShell(
    page3,
    PAGE.marginX,
    section6Top,
    CONTENT_WIDTH,
    section6Height,
    '6',
    'Pre-Filing Note',
    'How to interpret this package before final review and filing.',
    bodyFont,
    boldFont,
  );
  drawTextBlock(page3, closingNote, PAGE.marginX + 18, section6BodyTop, CONTENT_WIDTH - 36, bodyFont, 8.8, COLORS.inkSoft, 3.2);

  pages.forEach((page, idx) => drawFooter(page, idx + 1, pages.length, data.businessName || 'Business', data.taxYear, bodyFont, boldFont));

  pdfDoc.setTitle(sanitizePdfText(`MONIEZI Tax Prep Package Summary ${data.taxYear}`));
  pdfDoc.setAuthor('MONIEZI');
  pdfDoc.setCreator('MONIEZI Pro Finance');
  pdfDoc.setProducer('MONIEZI Pro Finance');
  pdfDoc.setSubject(sanitizePdfText(`Tax Prep Package Summary ${data.taxYear}`));
  return pdfDoc.save();
}
