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

const pageWidth = 595.28;
const pageHeight = 841.89;
const margin = 40;
const contentWidth = pageWidth - margin * 2;

const COLORS = {
  text: rgb(0.07, 0.11, 0.2),
  textSoft: rgb(0.37, 0.44, 0.54),
  line: rgb(0.86, 0.89, 0.94),
  panel: rgb(0.97, 0.98, 1),
  panelBorder: rgb(0.86, 0.9, 0.95),
  blue: rgb(0.16, 0.39, 0.89),
  blueTint: rgb(0.93, 0.96, 1),
  green: rgb(0.16, 0.73, 0.36),
  greenTint: rgb(0.88, 0.97, 0.91),
  red: rgb(0.89, 0.23, 0.23),
  redTint: rgb(1, 0.92, 0.92),
  yellow: rgb(0.92, 0.64, 0.14),
};


const sanitizePdfText = (value: unknown) => String(value ?? '')
  .normalize('NFKD')
  .replace(/[‐‑‒–—―]/g, '-')
  .replace(/[•·]/g, '-')
  .replace(/[“”]/g, '"')
  .replace(/[‘’]/g, "'")
  .replace(/…/g, '...')
  .replace(/\u00A0/g, ' ')
  .replace(/[^\x20-\x7E\n]/g, '');

const formatCurrency = (symbol: string, value: number) => `${symbol}${Number(value || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
const formatNumber = (value: number, decimals = 0) => Number(value || 0).toLocaleString(undefined, { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
const formatPercent = (value: number) => `${Number(value || 0).toLocaleString(undefined, { minimumFractionDigits: value % 1 === 0 ? 0 : 1, maximumFractionDigits: 1 })}%`;

const splitWords = (text: string, font: PDFFont, size: number, maxWidth: number) => {
  const lines: string[] = [];
  const safeText = sanitizePdfText(text);
  const words = safeText.split(/\s+/).filter(Boolean);
  if (!words.length) return [''];
  let current = words[0];
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

const drawTextBlock = (page: PDFPage, text: string, x: number, yTop: number, width: number, font: PDFFont, size: number, color = COLORS.textSoft, lineGap = 4) => {
  const lines = splitWords(text, font, size, width);
  const lineHeight = size + lineGap;
  let y = yTop - size;
  lines.forEach(line => {
    page.drawText(sanitizePdfText(line), { x, y, size, font, color });
    y -= lineHeight;
  });
  return y;
};

const drawLabelValueCard = (
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
  page.drawRoundedRectangle({ x, y: yTop - height, width, height, borderRadius: 16, borderWidth: 1, borderColor: COLORS.panelBorder, color: rgb(1, 1, 1) });
  page.drawText(sanitizePdfText(label).toUpperCase(), { x: x + 16, y: yTop - 30, size: 8.8, font: boldFont, color: COLORS.textSoft, characterSpacing: 1.6 });
  page.drawText(sanitizePdfText(value), { x: x + 16, y: yTop - 66, size: 22, font: boldFont, color: COLORS.text });
  drawTextBlock(page, note, x + 16, yTop - 86, width - 32, bodyFont, 7.5, COLORS.textSoft, 2.5);
};

const drawMetaCard = (page: PDFPage, x: number, yTop: number, width: number, height: number, label: string, value: string, bodyFont: PDFFont, boldFont: PDFFont) => {
  page.drawRoundedRectangle({ x, y: yTop - height, width, height, borderRadius: 14, borderWidth: 1, borderColor: COLORS.panelBorder, color: COLORS.blueTint });
  page.drawText(sanitizePdfText(label).toUpperCase(), { x: x + 14, y: yTop - 22, size: 7.8, font: boldFont, color: COLORS.textSoft, characterSpacing: 1.8 });
  const lines = splitWords(value, boldFont, 11, width - 28);
  let lineY = yTop - 52;
  lines.forEach(line => {
    page.drawText(sanitizePdfText(line), { x: x + 14, y: lineY, size: 11, font: boldFont, color: COLORS.text });
    lineY -= 15;
  });
};

const drawSectionShell = (page: PDFPage, x: number, yTop: number, width: number, height: number, sectionNo: string, title: string, subtitle: string, bodyFont: PDFFont, boldFont: PDFFont) => {
  page.drawRoundedRectangle({ x, y: yTop - height, width, height, borderRadius: 20, borderWidth: 1, borderColor: COLORS.panelBorder, color: rgb(1, 1, 1) });
  page.drawRoundedRectangle({ x, y: yTop - 86, width, height: 86, borderRadius: 20, color: COLORS.panel, borderColor: COLORS.panelBorder, borderWidth: 0 });
  page.drawText(sanitizePdfText(`SECTION ${sectionNo}`), { x: x + 18, y: yTop - 28, size: 8.8, font: boldFont, color: COLORS.blue, characterSpacing: 1.6 });
  page.drawText(sanitizePdfText(title), { x: x + 18, y: yTop - 58, size: 15, font: boldFont, color: COLORS.text });
  page.drawText(sanitizePdfText(subtitle), { x: x + 18, y: yTop - 76, size: 7.7, font: bodyFont, color: COLORS.textSoft });
  page.drawLine({ start: { x, y: yTop - 86 }, end: { x: x + width, y: yTop - 86 }, thickness: 1, color: COLORS.line });
  return yTop - 86;
};

const drawKeyValueRows = (page: PDFPage, x: number, yTop: number, width: number, rows: Array<{ key: string; note: string; value: string; emphasize?: boolean }>, bodyFont: PDFFont, boldFont: PDFFont) => {
  const rowHeight = 46;
  rows.forEach((row, idx) => {
    const rowTop = yTop - idx * rowHeight;
    if (idx > 0) {
      page.drawLine({ start: { x, y: rowTop }, end: { x: x + width, y: rowTop }, thickness: 1, color: COLORS.line });
    }
    if (row.emphasize) {
      page.drawRectangle({ x, y: rowTop - rowHeight, width, height: rowHeight, color: COLORS.blueTint });
    }
    page.drawText(sanitizePdfText(row.key), { x: x + 14, y: rowTop - 20, size: 8.6, font: boldFont, color: COLORS.text });
    page.drawText(sanitizePdfText(row.note), { x: x + 14, y: rowTop - 34, size: 7.2, font: row.emphasize ? boldFont : bodyFont, color: COLORS.textSoft });
    const safeValue = sanitizePdfText(row.value);
    const valueWidth = boldFont.widthOfTextAtSize(safeValue, 9.4);
    page.drawText(safeValue, { x: x + width - 14 - valueWidth, y: rowTop - 20, size: 9.4, font: boldFont, color: COLORS.text });
  });
};

const progressTone = (value: number) => {
  if (value >= 90) return { track: COLORS.greenTint, bar: COLORS.green, pill: COLORS.greenTint, text: rgb(0.09, 0.42, 0.21) };
  if (value >= 70) return { track: COLORS.blueTint, bar: COLORS.blue, pill: COLORS.blueTint, text: COLORS.blue };
  if (value >= 40) return { track: rgb(1, 0.96, 0.84), bar: COLORS.yellow, pill: rgb(1, 0.96, 0.84), text: rgb(0.64, 0.39, 0.03) };
  return { track: COLORS.redTint, bar: COLORS.red, pill: COLORS.redTint, text: COLORS.red };
};

const drawProgressRow = (page: PDFPage, x: number, yTop: number, width: number, label: string, detail: string, value: number, bodyFont: PDFFont, boldFont: PDFFont) => {
  const tone = progressTone(value);
  page.drawText(sanitizePdfText(label), { x, y: yTop - 16, size: 8.4, font: boldFont, color: COLORS.text });
  page.drawText(sanitizePdfText(detail), { x, y: yTop - 30, size: 7.2, font: bodyFont, color: COLORS.textSoft });
  page.drawRoundedRectangle({ x: x + width - 48, y: yTop - 24, width: 48, height: 20, borderRadius: 10, color: tone.pill });
  const pct = formatPercent(value);
  const pctWidth = boldFont.widthOfTextAtSize(sanitizePdfText(pct), 8.2);
  page.drawText(sanitizePdfText(pct), { x: x + width - 24 - pctWidth / 2, y: yTop - 17, size: 8.2, font: boldFont, color: tone.text });
  page.drawRoundedRectangle({ x, y: yTop - 46, width, height: 8, borderRadius: 4, color: tone.track });
  page.drawRoundedRectangle({ x, y: yTop - 46, width: Math.max(24, width * Math.max(0, Math.min(1, value / 100))), height: 8, borderRadius: 4, color: tone.bar });
};

const drawMiniStat = (page: PDFPage, x: number, yTop: number, width: number, height: number, label: string, value: string, note: string, bodyFont: PDFFont, boldFont: PDFFont) => {
  page.drawRoundedRectangle({ x, y: yTop - height, width, height, borderRadius: 16, borderWidth: 1, borderColor: COLORS.panelBorder, color: rgb(1,1,1) });
  page.drawText(sanitizePdfText(label).toUpperCase(), { x: x + 16, y: yTop - 24, size: 8.2, font: boldFont, color: COLORS.textSoft, characterSpacing: 1.5 });
  page.drawText(sanitizePdfText(value), { x: x + 16, y: yTop - 58, size: 20, font: boldFont, color: COLORS.text });
  drawTextBlock(page, note, x + 16, yTop - 76, width - 32, bodyFont, 7.1, COLORS.textSoft, 2.2);
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
  options?: { headerFontSize?: number; bodyFontSize?: number; rowHeight?: number; emptyStateItalic?: boolean },
) => {
  const headerHeight = 28;
  const rowHeight = options?.rowHeight ?? 34;
  const bodySize = options?.bodyFontSize ?? 8.6;
  const headerSize = options?.headerFontSize ?? 8.6;
  page.drawRectangle({ x, y: yTop - headerHeight, width, height: headerHeight, color: COLORS.panel });
  let colX = x;
  columns.forEach((col, idx) => {
    const headerLabel = sanitizePdfText(col.label);
    const textWidth = boldFont.widthOfTextAtSize(headerLabel, headerSize);
    const textX = col.align === 'right' ? colX + col.width - 12 - textWidth : colX + 12;
    page.drawText(headerLabel, { x: textX, y: yTop - 18, size: headerSize, font: boldFont, color: COLORS.textSoft, characterSpacing: idx === 0 ? 0.8 : 1.2 });
    colX += col.width;
  });
  page.drawLine({ start: { x, y: yTop - headerHeight }, end: { x: x + width, y: yTop - headerHeight }, thickness: 1, color: COLORS.line });
  rows.forEach((row, rowIndex) => {
    const top = yTop - headerHeight - rowIndex * rowHeight;
    if (rowIndex > 0) page.drawLine({ start: { x, y: top }, end: { x: x + width, y: top }, thickness: 1, color: COLORS.line });
    let cellX = x;
    row.forEach((cell, idx) => {
      const align = columns[idx]?.align ?? 'left';
      const font = options?.emptyStateItalic && row.length === 1 ? bodyFont : bodyFont;
      if (align === 'right') {
        const safeCell = sanitizePdfText(cell);
        const cellWidth = bodyFont.widthOfTextAtSize(safeCell, bodySize);
        page.drawText(safeCell, { x: cellX + columns[idx].width - 12 - cellWidth, y: top - 22, size: bodySize, font, color: COLORS.text });
      } else {
        page.drawText(sanitizePdfText(cell), { x: cellX + 12, y: top - 22, size: bodySize, font, color: COLORS.text });
      }
      cellX += columns[idx].width;
    });
  });
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
    // Keep the branded fonts when they load successfully.
    // Some installed mobile PWAs can fail to fetch/embed these assets reliably,
    // so we fall back to PDF standard fonts instead of aborting export.
    bodyFont = await pdfDoc.embedFont(reportRegularOtf, { subset: false });
    boldFont = await pdfDoc.embedFont(reportBoldOtf, { subset: false });
    kickerFont = await pdfDoc.embedFont(appRegularOtf, { subset: false });
  } catch (error) {
    console.warn('Tax Summary PDF custom fonts unavailable; falling back to standard PDF fonts.', error);
    bodyFont = await pdfDoc.embedFont(StandardFonts.Helvetica);
    boldFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
    kickerFont = boldFont;
  }

  const page1 = pdfDoc.addPage([pageWidth, pageHeight]);
  const page2 = pdfDoc.addPage([pageWidth, pageHeight]);
  const page3 = pdfDoc.addPage([pageWidth, pageHeight]);

  const summaryCards = [
    { label: 'Gross Business Income', value: formatCurrency(data.currencySymbol, data.totalIncome), note: `All income transactions recorded in MONIEZI for ${data.taxYear}.` },
    { label: 'Deductible Expenses', value: formatCurrency(data.currencySymbol, data.totalExpenses), note: `Across ${formatNumber(data.expenseItemsCount)} expense entries in this package.` },
    { label: 'Net Business Profit', value: formatCurrency(data.currencySymbol, data.netProfit), note: 'Before any final tax adjustments handled outside this report.' },
    { label: 'Mileage Deduction', value: formatCurrency(data.currencySymbol, data.mileageDeduction), note: `${formatNumber(data.totalMiles, 1)} business miles at ${data.currencySymbol}${formatNumber(data.mileageRate, 2)} per mile.` },
    { label: 'Ledger Transactions', value: formatNumber(data.ledgerTransactions), note: 'Income and expense entries included in this year-end package.' },
    { label: 'Linked Receipts', value: formatNumber(data.linkedReceipts), note: 'Receipt-backed expenses currently attached inside MONIEZI.' },
    { label: 'Expense Categories', value: formatNumber(data.expenseCategoriesCount), note: 'Distinct deduction buckets used in this tax year.' },
    { label: 'Top Expense Category', value: data.topExpenseCategoryName || '—', note: data.topExpenseCategoryName ? `${formatCurrency(data.currencySymbol, data.topExpenseCategoryAmount)} · ${formatPercent(data.topExpenseCategorySharePct)} of expenses` : 'No expense activity recorded for this period.' },
  ];

  // Page 1
  page1.drawText(sanitizePdfText('MONIEZI PRO FINANCE'), { x: margin, y: pageHeight - margin + 4, size: 11, font: kickerFont, color: COLORS.blue, characterSpacing: 2.1 });
  page1.drawText(sanitizePdfText('MONIEZI'), { x: margin, y: pageHeight - 92, size: 40, font: boldFont, color: COLORS.text });
  page1.drawText(sanitizePdfText(`Tax Prep Package Summary ${data.taxYear}`), { x: margin, y: pageHeight - 126, size: 22, font: boldFont, color: COLORS.text });
  drawTextBlock(page1, 'Year-end financial totals, documentation status, and mileage records prepared from your MONIEZI business data for accountant review and filing prep.', margin, pageHeight - 144, 320, bodyFont, 8.7, COLORS.textSoft, 4);

  drawMetaCard(page1, 418, pageHeight - 22, 137, 44, 'Business', data.businessName, bodyFont, boldFont);
  drawMetaCard(page1, 332, pageHeight - 78, 114, 54, 'Owner', data.ownerName, bodyFont, boldFont);
  drawMetaCard(page1, 452, pageHeight - 78, 103, 54, 'Generated', data.generatedAtLabel, bodyFont, boldFont);
  drawMetaCard(page1, 392, pageHeight - 144, 163, 46, 'Reporting Period', data.reportingPeriodLabel, bodyFont, boldFont);

  const cardGap = 12;
  const cardWidth = (contentWidth - cardGap * 3) / 4;
  const cardHeight = 95;
  const cardsTop = pageHeight - 208;
  summaryCards.forEach((card, idx) => {
    const col = idx % 4;
    const row = Math.floor(idx / 4);
    drawLabelValueCard(page1, margin + col * (cardWidth + cardGap), cardsTop - row * (cardHeight + cardGap), cardWidth, cardHeight, card.label, card.value, card.note, bodyFont, boldFont);
  });

  const section1Top = pageHeight - 422;
  drawSectionShell(page1, margin, section1Top, contentWidth, 408, '1', 'Tax-Ready Financial Snapshot', 'Core totals your accountant typically needs first, presented in one clean year-end view.', bodyFont, boldFont);
  drawKeyValueRows(page1, margin, section1Top - 86, contentWidth, [
    { key: 'Gross Business Income', note: 'Total recorded income transactions for the selected tax year.', value: formatCurrency(data.currencySymbol, data.totalIncome) },
    { key: 'Deductible Business Expenses', note: 'Total expense entries tracked in MONIEZI before any external adjustments.', value: formatCurrency(data.currencySymbol, data.totalExpenses) },
    { key: 'Net Business Profit', note: 'Income less recorded expenses for the selected period.', value: formatCurrency(data.currencySymbol, data.netProfit), emphasize: true },
    { key: 'Business Mileage Logged', note: `${formatNumber(data.completeMileageCount)} complete trip ${data.completeMileageCount === 1 ? 'entry' : 'entries'} captured in the mileage log.`, value: `${formatNumber(data.totalMiles, 1)} mi` },
    { key: 'Standard Mileage Rate Used', note: 'Configured inside MONIEZI settings for the selected export.', value: `${data.currencySymbol}${formatNumber(data.mileageRate, 2)} / mi` },
    { key: 'Estimated Mileage Deduction', note: 'Computed from logged business miles using the configured rate.', value: formatCurrency(data.currencySymbol, data.mileageDeduction) },
    { key: 'Reporting Period', note: 'Earliest to latest record included in this export package.', value: data.reportingPeriodLabel },
  ], bodyFont, boldFont);

  // Page 2
  page2.drawText(sanitizePdfText('MONIEZI TAX PREP PACKAGE'), { x: margin, y: pageHeight - margin + 4, size: 11, font: kickerFont, color: COLORS.blue, characterSpacing: 2.1 });
  page2.drawText(sanitizePdfText('Documentation Status & Expense Breakdown'), { x: margin, y: pageHeight - 92, size: 25, font: boldFont, color: COLORS.text });
  drawTextBlock(page2, 'A fixed summary page showing record completeness plus the highest-impact deduction categories for the year.', margin, pageHeight - 124, 340, bodyFont, 8.8, COLORS.textSoft, 4);
  drawMetaCard(page2, 458, pageHeight - 22, 97, 44, 'Tax Year', data.taxYear, bodyFont, boldFont);
  drawMetaCard(page2, 436, pageHeight - 78, 119, 44, 'Expense Categories', formatNumber(data.expenseCategoriesCount), bodyFont, boldFont);

  const section2Top = pageHeight - 150;
  drawSectionShell(page2, margin, section2Top, contentWidth, 310, '2', 'Audit Readiness & Documentation Status', 'A quick view of how complete and organized your records look before filing.', bodyFont, boldFont);
  const progressBaseY = section2Top - 104;
  drawProgressRow(page2, margin + 18, progressBaseY, contentWidth - 36, 'Receipt Coverage', `${formatNumber(data.linkedReceipts)} linked receipts across ${formatNumber(data.expenseItemsCount)} deductible expense items.`, data.receiptCoveragePct, bodyFont, boldFont);
  drawProgressRow(page2, margin + 18, progressBaseY - 74, contentWidth - 36, 'Expense Review Status', `${formatNumber(data.reviewedExpenseCount)} reviewed · ${formatNumber(data.pendingReviewCount)} pending review.`, data.reviewCoveragePct, bodyFont, boldFont);
  drawProgressRow(page2, margin + 18, progressBaseY - 148, contentWidth - 36, 'Mileage Log Completeness', `${formatNumber(data.completeMileageCount)} complete trip entries recorded for ${formatNumber(data.totalMiles, 1)} business miles.`, data.mileageCompletionPct, bodyFont, boldFont);
  const miniTop = section2Top - 222;
  const miniGap = 14;
  const miniWidth = (contentWidth - miniGap * 2 - 36) / 3;
  drawMiniStat(page2, margin + 18, miniTop, miniWidth, 88, 'Package Coverage', formatNumber(data.ledgerTransactions), 'Total ledger transactions included in this tax-prep package export.', bodyFont, boldFont);
  drawMiniStat(page2, margin + 18 + miniWidth + miniGap, miniTop, miniWidth, 88, 'Items Requiring Attention', formatNumber(data.itemsRequiringAttention), 'Headline open items across receipts, review status, categorization, and mileage completeness.', bodyFont, boldFont);
  drawMiniStat(page2, margin + 18 + (miniWidth + miniGap) * 2, miniTop, miniWidth, 88, 'Prepared Privately', '100%', 'Generated directly from your MONIEZI records for local export and review.', bodyFont, boldFont);

  const section3Top = pageHeight - 438;
  drawSectionShell(page2, margin, section3Top, contentWidth, 368, '3', 'Deductible Expense Breakdown', 'Top expense categories by dollar amount, including category share and receipt-backed count.', bodyFont, boldFont);
  drawTable(page2, margin, section3Top - 86, contentWidth, [
    { label: 'Expense Category', width: 235 },
    { label: 'Amount', width: 120, align: 'right' },
    { label: 'Share', width: 80, align: 'right' },
    { label: 'Receipts', width: 80, align: 'right' },
  ],
  data.expenseRows.length
    ? data.expenseRows.map(row => [row.name, formatCurrency(data.currencySymbol, row.amount), formatPercent(row.sharePct), `${row.linked}/${row.count}`])
    : [['No deductible expenses were recorded for this tax year.', '', '', '']],
  bodyFont, boldFont, { rowHeight: 36, bodyFontSize: 8.8, headerFontSize: 8.4 });

  // Page 3
  page3.drawText(sanitizePdfText('MONIEZI TAX PREP PACKAGE'), { x: margin, y: pageHeight - margin + 4, size: 11, font: kickerFont, color: COLORS.blue, characterSpacing: 2.1 });
  page3.drawText(sanitizePdfText('Mileage & Filing Checks'), { x: margin, y: pageHeight - 92, size: 25, font: boldFont, color: COLORS.text });
  drawTextBlock(page3, 'A compact closing page focused on quarter-level mileage totals, filing cleanup items, and final handoff guidance.', margin, pageHeight - 124, 360, bodyFont, 8.8, COLORS.textSoft, 4);
  drawMetaCard(page3, 370, pageHeight - 34, 185, 52, 'Reporting Period', data.reportingPeriodLabel, bodyFont, boldFont);

  const splitGap = 16;
  const splitWidth = (contentWidth - splitGap) / 2;
  const section4Top = pageHeight - 168;
  drawSectionShell(page3, margin, section4Top, splitWidth, 270, '4', 'Mileage Log Summary', 'Quarter-by-quarter view of recorded trips, miles, and estimated deduction.', bodyFont, boldFont);
  if (data.hasMileageRows) {
    drawTable(page3, margin, section4Top - 86, splitWidth, [
      { label: 'Quarter', width: 88 },
      { label: 'Trips', width: 54, align: 'right' },
      { label: 'Miles', width: 86, align: 'right' },
      { label: 'Deduction', width: splitWidth - 88 - 54 - 86, align: 'right' },
    ], data.quarterlyMileage.map(row => [row.quarter, formatNumber(row.trips), formatNumber(row.miles, 1), formatCurrency(data.currencySymbol, row.deduction)]), bodyFont, boldFont, { headerFontSize: 7.6, bodyFontSize: 8.6, rowHeight: 36 });
  } else {
    page3.drawText(sanitizePdfText('No mileage trips were recorded for this tax year.'), { x: margin + 18, y: section4Top - 132, size: 10.5, font: bodyFont, color: COLORS.textSoft });
  }

  drawSectionShell(page3, margin + splitWidth + splitGap, section4Top, splitWidth, 188, '5', 'Attention Items Before Filing', 'The items below help explain where additional cleanup or support documents may still be needed.', bodyFont, boldFont);
  let noteY = section4Top - 112;
  data.attentionItems.slice(0, 4).forEach(item => {
    const wrapped = splitWords(item, bodyFont, 9.2, splitWidth - 48);
    wrapped.forEach((line, idx) => {
      if (idx === 0) page3.drawCircle({ x: margin + splitWidth + splitGap + 18, y: noteY + 4, size: 2.6, color: COLORS.blue });
      page3.drawText(sanitizePdfText(line), { x: margin + splitWidth + splitGap + 28, y: noteY, size: 9.2, font: bodyFont, color: COLORS.text });
      noteY -= 13;
    });
    noteY -= 12;
  });

  const closingTop = pageHeight - 428;
  page3.drawRoundedRectangle({ x: margin, y: closingTop - 110, width: contentWidth, height: 110, borderRadius: 18, borderWidth: 1, borderColor: COLORS.panelBorder, color: COLORS.panel });
  page3.drawText(sanitizePdfText('Pre-Filing Note'), { x: margin + 18, y: closingTop - 28, size: 12.5, font: boldFont, color: COLORS.text });
  drawTextBlock(page3, 'MONIEZI organized this package from your recorded ledger entries, linked receipt attachments, and mileage logs for the selected tax year. The totals here are designed to make the value of your records immediately clear: what you earned, what you spent, how well expenses are documented, and what should be addressed before filing. Final tax treatment, classification decisions, and any required adjustments should still be reviewed with your tax professional.', margin + 18, closingTop - 44, contentWidth - 36, bodyFont, 9.2, COLORS.textSoft, 4.2);

  page3.drawLine({ start: { x: margin, y: margin + 18 }, end: { x: pageWidth - margin, y: margin + 18 }, thickness: 1, color: COLORS.line });
  page3.drawText(sanitizePdfText('MONIEZI Pro Finance · Generated privately from your local business records.'), { x: margin, y: margin, size: 9.2, font: boldFont, color: COLORS.textSoft });
  const footerRight = `${data.businessName} · Tax Year ${data.taxYear}`;
  const safeFooterRight = sanitizePdfText(footerRight);
  const footerWidth = bodyFont.widthOfTextAtSize(safeFooterRight, 9.2);
  page3.drawText(safeFooterRight, { x: pageWidth - margin - footerWidth, y: margin, size: 9.2, font: bodyFont, color: COLORS.textSoft });

  pdfDoc.setTitle(sanitizePdfText(`MONIEZI Tax Prep Package Summary ${data.taxYear}`));
  pdfDoc.setAuthor('MONIEZI');
  pdfDoc.setCreator('MONIEZI Pro Finance');
  pdfDoc.setProducer('MONIEZI Pro Finance');
  pdfDoc.setSubject(sanitizePdfText(`Tax Prep Package Summary ${data.taxYear}`));
  return pdfDoc.save();
}
