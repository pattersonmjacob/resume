function onOpen() {
  DocumentApp.getUi().createMenu('Resume Tools')
    .addItem('Open Resume Sidebar', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Update Resume Section')
    .setWidth(400);
  DocumentApp.getUi().showSidebar(html);
}

function replaceSectionAndLocationAndTitles(headerText, newContent, relocationText, strategyOpsTitle, strategyInsightsTitle, financeOpsTitle) {
  replaceSection(headerText, newContent);
  updateRelocationLine(relocationText);

  updateRoleTitle([
    "Strategy & Operations (Senior Manager)",
    "Strategic Operations (Senior Manager)",
    "Strategy & Operations - Program Management Office (Senior Manager)"
  ], strategyOpsTitle);

  updateRoleTitle([
    "Strategy & Insights (Global Program Lead)",
    "Strategy & Insights - Data & Analytics (Global Program Lead)"
  ], strategyInsightsTitle);

  updateRoleTitle([
    "Finance & Business Operations (Senior Analyst, Analyst II, & Analyst I)",
    "Finance & Commercial Operations (Senior Analyst, Analyst II, & Analyst I)",
    "Finance & Business Operations - Revenue (Senior Analyst, Analyst II, & Analyst I)",
    "Finance & Business Operations - Sales (Senior Analyst, Analyst II, & Analyst I)",
    "Finance & Commercial Operations - Revenue (Senior Analyst, Analyst II, & Analyst I)",
    "Finance & Commercial Operations - Deal Desk (Senior Analyst, Analyst II, & Analyst I)"
  ], financeOpsTitle);

  return true;
}
function updateRelocationLine(newLocationText) {
  const body = DocumentApp.getActiveDocument().getBody();
  const paragraphs = body.getParagraphs();

  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    const text = para.getText();

    // Only act on the contact line
    if (text.includes('|') && (text.includes('Open to Relocation') || text.includes('Remote – US'))) {
      const emojiAnchor = '🌍 ';
      const emojiIndex = text.indexOf(emojiAnchor);
      if (emojiIndex === -1) continue;

      const preserved = text.substring(0, emojiIndex + emojiAnchor.length);
      const updatedLine = preserved + newLocationText.trim();

      // Apply update without altering formatting
      const textElement = para.editAsText();
      const preservedLength = preserved.length;

      // Only update the part after the 🌍 emoji
      textElement.deleteText(preservedLength, text.length - 1);
      textElement.insertText(preservedLength, newLocationText.trim());

      return;
    }
  }
}


function updateRoleTitle(knownTitles, selectedTitle) {
  const body = DocumentApp.getActiveDocument().getBody();
  const paragraphs = body.getParagraphs();

  for (let i = 0; i < paragraphs.length; i++) {
    const fullText = paragraphs[i].getText().trim();
    const rolePart = fullText.split(/\t| {2,}/)[0].trim();
    const normalizedPara = normalizeText(rolePart);

    for (let j = 0; j < knownTitles.length; j++) {
      const expected = knownTitles[j];
      const normalizedExpected = normalizeText(expected);

      if (normalizedPara === normalizedExpected) {
        const para = paragraphs[i].editAsText();
        const newText = fullText.replace(rolePart, selectedTitle);
        para.setText(newText);

        const endIdx = selectedTitle.length;
        para.setFontFamily(0, endIdx - 1, 'Times New Roman');
        para.setFontSize(0, endIdx - 1, 10.5);
        para.setItalic(0, endIdx - 1, true);
        para.setBold(0, endIdx - 1, false);

        return;
      }
    }
  }

  throw new Error("Role title not found. Double-check document formatting.");
}

function normalizeText(text) {
  if (!text || typeof text !== 'string') return '';
  return text
    .toLowerCase()
    .replace(/[–—]/g, '-')
    .replace(/&amp;/g, '&')
    .replace(/\s+/g, ' ')
    .replace(/[^\w\s\-\(\)&]/g, '')
    .trim();
}

function replaceSection(headerText, newContent) {
  const body = DocumentApp.getActiveDocument().getBody();
  const paragraphs = body.getParagraphs();
  const targetHeader = headerText.trim().toLowerCase();
  let startIndex = -1;
  let endIndex = paragraphs.length;

  for (let i = 0; i < paragraphs.length; i++) {
    const text = paragraphs[i].getText().trim().toLowerCase();
    if (text === targetHeader) {
      startIndex = i + 1;
      break;
    }
  }

  if (startIndex === -1) {
    throw new Error(`Section header "${headerText}" not found.`);
  }

  for (let i = startIndex; i < paragraphs.length; i++) {
    const text = paragraphs[i].getText().trim();
    const isHeading = paragraphs[i].getHeading() !== DocumentApp.ParagraphHeading.NORMAL;

    if (isHeading && text !== '') {
      endIndex = i;
      break;
    }

    if (text === '' && i + 1 < paragraphs.length) {
      const nextPara = paragraphs[i + 1];
      if (nextPara.getHeading() !== DocumentApp.ParagraphHeading.NORMAL && nextPara.getText().trim() !== '') {
        endIndex = i;
        break;
      }
    }
  }

  let targetPara = paragraphs[startIndex];
  if (!targetPara) {
    targetPara = body.insertParagraph(startIndex, '');
  }

  const textElement = targetPara.editAsText();
  const sanitized = newContent
    .split('\n')
    .map(line => line.trim())
    .filter(line => line.length > 0)
    .join(' ');
  textElement.setText(sanitized);
  return true;
}

function scanForFormattingIssues() {
  const body = DocumentApp.getActiveDocument().getBody();
  const paragraphs = body.getParagraphs();
  const issues = [];

  let inTechSection = false;
  let techLines = [];

  for (let i = 0; i < paragraphs.length; i++) {
    const text = paragraphs[i].getText().trim();

    if (!inTechSection && text.toLowerCase() === 'technologies & competencies') {
      inTechSection = true;
      continue;
    }

    if (inTechSection) {
      if (
        paragraphs[i].getHeading() !== DocumentApp.ParagraphHeading.NORMAL ||
        text === ''
      ) break;

      techLines.push({ lineNum: i + 1, content: text });
    }
  }

  const fullText = techLines.map(t => t.content).join(' ');

  // Check for duplicate skills (case-insensitive)
  const tokens = fullText
    .split('•')
    .map(t => t.trim().toLowerCase())
    .filter(t => t.length > 0);

  const seen = new Set();
  const duplicates = new Set();
  for (const token of tokens) {
    if (seen.has(token)) duplicates.add(token);
    else seen.add(token);
  }

  if (duplicates.size > 0) {
    issues.push(`TECHNOLOGIES & COMPETENCIES: Repeated skills — ${Array.from(duplicates).join(', ')}`);
  }

  // Check for back-to-back repeated words (case-insensitive)
  const words = fullText.split(/\s+/);
  for (let i = 0; i < words.length - 1; i++) {
    if (words[i].toLowerCase() === words[i + 1].toLowerCase()) {
      issues.push(`TECHNOLOGIES & COMPETENCIES: Repeated word — "${words[i]} ${words[i + 1]}"`);
    }
  }

  // Per-line formatting checks
  for (const { lineNum, content } of techLines) {
    if (content.includes('  ')) {
      issues.push(`Line ${lineNum}: Contains double spaces`);
    }

    const bulletIssues = [];
    const improperBullets = /[^ ]•|•[^ ]/g;
    let match;
    while ((match = improperBullets.exec(content)) !== null) {
      bulletIssues.push(`improper bullet spacing at index ${match.index}`);
    }

    if (bulletIssues.length) {
      issues.push(`Line ${lineNum}: ${bulletIssues.join(', ')}`);
    }
  }

  return issues;
}
