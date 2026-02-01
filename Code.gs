/** @OnlyCurrentDoc */

// ========== 지능형 서식 엔진 (STYLE_CONFIG) ==========
// 학회별 규정에 유연히 대응하는 변수 기반 서식 엔진.
// STYLE_CONFIG = DocumentProperties에 저장된 런타임 설정 (getEffectiveConfig()로 조회).

/** 스타일 프리셋: APA 7th, MLA, Custom 시 기본값으로 사용 */
var STYLE_PRESETS = {
  APA: {
    italicsTitle: true,
    italicsJournal: true,
    yearBracket: '()',
    authorSeparator: ', ',
    etAlThreshold: 3,
    etAlEn: 'et al.',
    etAlKo: '외',
    headerFontSize: 14,
    headerAlignment: 'CENTER',
    lineSpacing: 12,
    hangingIndent: 20,
    indentFirstLine: 0,
    korItalicsJournal: false,
    usePagePrefix: true,
    korUsePagePrefix: false,
    bibliographyUseEtAl: false
  },
  MLA: {
    italicsTitle: true,
    italicsJournal: true,
    yearBracket: 'none',
    authorSeparator: ', ',
    etAlThreshold: 3,
    etAlEn: 'et al.',
    etAlKo: '외',
    headerFontSize: 14,
    headerAlignment: 'CENTER',
    lineSpacing: 12,
    hangingIndent: 20,
    indentFirstLine: 0,
    korItalicsJournal: false,
    usePagePrefix: true,
    korUsePagePrefix: false,
    bibliographyUseEtAl: false
  }
};

/** 현재 문서에 적용된 STYLE_CONFIG 객체 반환 (저장값 없으면 APA 프리셋) */
function getEffectiveConfig() {
  var props = PropertiesService.getDocumentProperties();
  var saved = props.getProperty('STYLE_CONFIG');
  if (!saved) return Object.assign({}, STYLE_PRESETS.APA);
  try {
    var parsed = JSON.parse(saved);
    var base = parsed.preset && STYLE_PRESETS[parsed.preset] ? STYLE_PRESETS[parsed.preset] : STYLE_PRESETS.APA;
    return Object.assign({}, base, parsed.config || {});
  } catch (e) {
    return Object.assign({}, STYLE_PRESETS.APA);
  }
}

function getStyleConfig() {
  try {
    var props = PropertiesService.getDocumentProperties();
    var saved = props.getProperty('STYLE_CONFIG');
    if (!saved) return JSON.stringify({ ok: true, config: STYLE_PRESETS.APA, preset: 'APA' });
    var parsed = JSON.parse(saved);
    var base = (parsed.preset && STYLE_PRESETS[parsed.preset]) ? STYLE_PRESETS[parsed.preset] : STYLE_PRESETS.APA;
    var config = Object.assign({}, base, parsed.config || {});
    return JSON.stringify({ ok: true, config: config, preset: parsed.preset || 'APA' });
  } catch (e) {
    return JSON.stringify({ ok: true, config: STYLE_PRESETS.APA, preset: 'APA' });
  }
}

function saveStyleConfig(configObj) {
  try {
    var props = PropertiesService.getDocumentProperties();
    var cfg = configObj.config || configObj;
    var preset = configObj.preset || 'Custom';
    if (preset === 'APA' || preset === 'MLA') {
      cfg = Object.assign({}, STYLE_PRESETS[preset], cfg);
    }
    props.setProperty('STYLE_CONFIG', JSON.stringify({ config: cfg, preset: preset }));
    return JSON.stringify({ ok: true });
  } catch (e) {
    return JSON.stringify({ ok: false, error: (e && e.message) ? e.message : String(e) });
  }
}

function isKoreanCitation(authorStr) {
  if (!authorStr || typeof authorStr !== 'string') return false;
  return /[\uac00-\ud7a3\u4e00-\u9fff]/.test(authorStr);
}

/** 문장 내 인용(In-text)용 저자 포맷: 국문은 '외', 영문은 2인 '성 & 성', 3인 이상 '첫 저자 성 et al.' 강제 */
function formatInTextCitation(authorRaw, year, style) {
  var isKo = isKoreanCitation(authorRaw || '');
  var author = formatAuthorsForInText(authorRaw || '', isKo);
  var s = (style || 'APA').toUpperCase();
  return s === 'APA' ? '(' + author + ', ' + year + ')' : '(' + author + ' ' + year + ')';
}

/** 저자 문자열을 지능적으로 분리하는 함수 */
function smartSplitAuthors(authorStr, isKo) {
  if (!authorStr) return [];
  
  // 1. 세미콜론(;)이 포함되어 있다면, 사용자가 명확히 구분한 것으로 간주하고 세미콜론으로만 자릅니다.
  if (authorStr.indexOf(';') >= 0) {
    return authorStr.split(';').map(function(s) { return s.trim(); }).filter(function(s) { return s; });
  }
  
  // 2. 세미콜론이 없고 국문 문헌인 경우에만 쉼표(,)를 구분자로 인정합니다.
  if (isKo) {
    return authorStr.split(',').map(function(s) { return s.trim(); }).filter(function(s) { return s; });
  }
  
  // 3. 세미콜론이 없는 영문 문헌의 경우, 쉼표가 있더라도 일단 '한 명의 이름'으로 간주하는 것이 안전합니다.
  // (예: Smith, John -> 1명 / 만약 2명이라면 Smith; Doe 처럼 세미콜론 사용 권장)
  return [authorStr.trim()];
}

/** 위 함수를 적용한 인용 축약 로직 */
function formatAuthorsForInText(authorStr, isKo) {
  var parts = smartSplitAuthors(authorStr, isKo);
  if (parts.length === 0) return 'Unknown';

  function getSurname(name, isKo) {
    var p = (name || '').trim();
    if (isKo) return p; // 국문은 전체 이름 사용
    // 영문: Lastname, Firstname 형식이면 쉼표 앞이 성. 아니면 공백 뒤가 성.
    var commaIndex = p.indexOf(',');
    if (commaIndex >= 0) return p.substring(0, commaIndex).trim();
    var spaceParts = p.split(' ');
    return spaceParts[spaceParts.length - 1];
  }

  if (isKo) {
    if (parts.length >= 3) return parts[0] + ' 외';
    if (parts.length === 2) return parts[0] + '·' + parts[1];
    return parts[0];
  } else {
    if (parts.length >= 3) return getSurname(parts[0], false) + ' et al.';
    if (parts.length === 2) return getSurname(parts[0], false) + ' & ' + getSurname(parts[1], false);
    return getSurname(parts[0], false);
  }
}

/** 사이드바용: 문장 내 인용 문자열 반환 (국문=외, 영문=et al.) */
function getInTextCitationString(authorRaw, year, style) {
  return formatInTextCitation(authorRaw || '', year || '', style || 'APA');
}

/**
 * 통째로 붙여넣은 참고문헌 텍스트를 분석하여 본문 인용구를 제안.
 * 스타일 앵커(큰따옴표=MLA, 괄호 연도=APA) → 저자 블록 분리 → (?<=\.), 구분자로 개별 저자 식별 → formatAuthorsForInText로 축약형 생성.
 * @param {string} fullText - 전체 참고문헌 정보(Full) 텍스트
 * @return {string} JSON. { ok: true, citationText: "(저자, 연도)", year: "연도" } 또는 { ok: false, error: "메시지" }
 */
function parseAndSuggestCitation(fullText) {
  try {
    if (!fullText || typeof fullText !== 'string') {
      return JSON.stringify({ ok: false, error: '입력된 텍스트가 없습니다.' });
    }
    var text = fullText.trim();
    if (!text) {
      return JSON.stringify({ ok: false, error: '입력된 텍스트가 없습니다.' });
    }

    // 1) 스타일 앵커 탐색: 큰따옴표 → MLA(저자=따옴표 직전), 괄호 4자리 연도 → APA(저자=괄호 직전)
    var authorBlock = '';
    var quoteIdx = text.indexOf('"');
    var apaMatch = text.match(/\((19|20)\d{2}\)/);
    if (quoteIdx >= 0) {
      authorBlock = text.substring(0, quoteIdx).trim();
    } else if (apaMatch) {
      var parenIdx = text.indexOf(apaMatch[0]);
      authorBlock = text.substring(0, parenIdx).trim();
    } else {
      authorBlock = text;
    }

    // 연도: 텍스트 내에서 가장 먼저 발견되는 4자리 숫자(19XX/20XX)
    var yearMatch = text.match(/(19|20)\d{2}/);
    var year = yearMatch ? yearMatch[0] : '';

    // 2) 저자 블록 내 개별 저자 식별: (?<=\.), 는 저자 구분자. 마침표+알파벳/공백은 이니셜로 보호
    var authorParts = [];
    try {
      authorParts = authorBlock.split(/(?<=\.),\s*/);
    } catch (e) {
      authorParts = authorBlock.split(/\.\s*,\s*/);
    }
    authorParts = authorParts.map(function(s) { return s.trim(); }).filter(function(s) { return s; });
    if (authorParts.length > 0) {
      authorParts[authorParts.length - 1] = authorParts[authorParts.length - 1].replace(/\s*\(\s*(19|20)\d{2}\s*\)\s*$/, '').trim();
    }
    var authorJoined = authorParts.length > 0 ? authorParts.join('; ') : authorBlock;
    if (!authorJoined) authorJoined = authorBlock;

    // 3) formatAuthorsForInText에 전달하여 축약형 생성 후 (저자, 연도) 형태로 조합
    var isKo = isKoreanCitation(authorJoined);
    var formattedAuthor = formatAuthorsForInText(authorJoined, isKo);
    var citationText = year ? '(' + formattedAuthor + ', ' + year + ')' : '(' + formattedAuthor + ')';

    return JSON.stringify({ ok: true, citationText: citationText, year: year });
  } catch (e) {
    var errMsg = (e && e.message) ? e.message : String(e);
    return JSON.stringify({ ok: false, error: '파싱 중 오류가 발생했습니다. 본문 인용구를 수동으로 입력해 주세요. (' + errMsg + ')' });
  }
}

/**
 * 저자 문자열을 STYLE_CONFIG에 따라 포맷 (국문 '외', 영문 'et al.').
 * @param {boolean} isBibliography - true면 참고문헌 목록용으로 etAlThreshold 무시·전체 명단 반환 (bibliographyUseEtAl가 true일 때만 et al./외 적용)
 */
function formatAuthorsWithConfig(authorStr, cfg, isKo, isBibliography) {
  if (!authorStr) return 'Unknown';
  var parts = authorStr.split(';').map(function(s) { return s.trim(); }).filter(function(s) { return s; });
  if (parts.length === 0) return 'Unknown';
  var sep = (cfg && cfg.authorSeparator !== undefined) ? cfg.authorSeparator : ', ';
  var etAl = isKo ? (cfg && cfg.etAlKo) ? cfg.etAlKo : '외' : (cfg && cfg.etAlEn) ? cfg.etAlEn : 'et al.';
  var th = (cfg && parseInt(cfg.etAlThreshold, 10)) ? parseInt(cfg.etAlThreshold, 10) : 3;
  if (isBibliography && !(cfg && cfg.bibliographyUseEtAl === true)) return parts.join(sep);
  if (parts.length <= th) return parts.join(sep);
  return parts[0] + ' ' + etAl;
}

/**
 * 문헌 항목을 STYLE_CONFIG에 따라 포맷팅 (국문/영문 이원화, 리치 텍스트 범위 계산).
 * 저자 문자열은 authorRaw(원본) 또는 author를 formatAuthorsWithConfig로 포맷하여 사용.
 * @return {{ text: string, ranges: Array<{ start: number, end: number, format: string }> }}
 */
function formatCitationEntry(item, cfg) {
  if (!item) return { text: '', ranges: [] };
  cfg = cfg || getEffectiveConfig();
  var ranges = [];

  if (item.mode === 'paste') {
    return { text: (item.fullText || '').trim(), ranges: [] };
  }

  var authorRaw = (item.authorRaw || item.author || '').trim();
  var isKo = isKoreanCitation(authorRaw);
  var author = formatAuthorsWithConfig(authorRaw, cfg, isKo, true);

  var title = (item.title || '').trim();
  var year = (item.secondVal || '').trim();
  var publisher = (item.publisher || '').trim();
  var volume = (item.volume || '').trim();
  var issue = (item.issue || '').trim();
  var page = (item.page || '').trim();
  var styleName = (item.style || 'APA').toUpperCase();

  var italicsTitle = cfg.italicsTitle === true;
  var italicsJournal = cfg.italicsJournal === true;
  if (isKo && cfg.korItalicsJournal === false) italicsJournal = false;
  var usePagePrefix = cfg.usePagePrefix !== false;
  var usePpForKor = (cfg.korUsePagePrefix === true) || (cfg.useKorPp === true);
  var korUsePagePrefix = isKo ? usePpForKor : usePagePrefix;
  var pagePrefix = (styleName === 'MLA' && (isKo ? korUsePagePrefix : usePagePrefix)) ? (cfg.pagePrefix !== undefined ? cfg.pagePrefix : 'pp. ') : '';

  var yearPart = '';
  if (cfg.yearBracket === '()') yearPart = ' (' + year + ')';
  else if (cfg.yearBracket === '[]') yearPart = ' [' + year + ']';
  else if (year) yearPart = ' ' + year;

  var fullText = '';
  if (styleName === 'MLA') {
    fullText = author + '. "' + title + '." ';
    var mlaTitleStart = author.length + 3;
    var mlaTitleEnd = mlaTitleStart + title.length;
    if (italicsTitle && title.length > 0) {
      ranges.push({ start: mlaTitleStart, end: mlaTitleEnd, format: 'italics' });
    }
    var pos = fullText.length;
    if (publisher) {
      var mlaPubStart = pos;
      fullText += publisher + ', ';
      if (italicsJournal && publisher.length > 0) {
        ranges.push({ start: mlaPubStart, end: mlaPubStart + publisher.length, format: 'italics' });
      }
    }
    if (volume) fullText += 'vol. ' + volume + (issue ? ', no. ' + issue : '') + (page ? ', ' + pagePrefix + page : '') + '. ';
    else if (issue) fullText += 'no. ' + issue + (page ? ', ' + pagePrefix + page : '') + '. ';
    else if (page) fullText += pagePrefix + page + '. ';
    fullText += year + '.';
  } else {
    fullText = author + yearPart + '. ' + title + '. ';
    var titleStart = author.length + yearPart.length + 2;
    var titleEnd = titleStart + title.length;
    if (italicsTitle && title.length > 0) {
      ranges.push({ start: titleStart, end: titleEnd, format: 'italics' });
    }
    var pos = fullText.length;
    if (publisher) {
      var pubStart = pos;
      fullText += publisher;
      if (italicsJournal && publisher.length > 0) {
        ranges.push({ start: pubStart, end: pubStart + publisher.length, format: 'italics' });
      }
      if (volume || issue || page) {
        var volIssue = volume ? (volume + (issue ? '(' + issue + ')' : '')) : (issue ? '(' + issue + ')' : '');
        if (volIssue) fullText += ', ' + volIssue;
        if (page) fullText += ', ' + (isKo && !korUsePagePrefix ? '' : pagePrefix) + page + '.';
        else if (volIssue) fullText += '.';
      } else {
        fullText += '.';
      }
    } else {
      if (volume || issue || page) {
        var volIssue2 = volume ? (volume + (issue ? '(' + issue + ')' : '')) : (issue ? '(' + issue + ')' : '');
        if (volIssue2) fullText += volIssue2;
        if (page) fullText += (volIssue2 ? ', ' : '') + (isKo && !korUsePagePrefix ? '' : pagePrefix) + page + '.';
        else if (volIssue2) fullText += '.';
      }
    }
  }

  return { text: fullText.trim(), ranges: ranges };
}

function onOpen() {
  DocumentApp.getUi().createMenu('인용 관리 메뉴')
      .addItem('사이드바 열기', 'showSidebar')
      .addSeparator()
      .addItem('데이터 전체 초기화', 'clearCitations')
      .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('CitationManager').setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}

function clearCitations() {
  PropertiesService.getDocumentProperties().deleteProperty('CITATION_LIST');
}

function saveCitationOnly(data) {
  try {
    if (!data || typeof data !== 'object') return JSON.stringify({ ok: false, error: '저장할 데이터가 없습니다.' });
    const props = PropertiesService.getDocumentProperties();
    if (data.korUsePagePrefix !== undefined) {
      var saved = props.getProperty('STYLE_CONFIG');
      var preset = 'APA';
      if (saved) try { var p = JSON.parse(saved); preset = p.preset || 'APA'; } catch (e) {}
      var cfg = getEffectiveConfig();
      cfg.korUsePagePrefix = data.korUsePagePrefix === true;
      saveStyleConfig({ config: cfg, preset: preset });
      delete data.korUsePagePrefix;
    }
    let citationList = [];
    try {
      const citations = props.getProperty('CITATION_LIST');
      if (citations) citationList = JSON.parse(citations);
    } catch (e) {
      citationList = [];
    }
    if (!Array.isArray(citationList)) citationList = [];
    citationList.push(data);
    props.setProperty('CITATION_LIST', JSON.stringify(citationList));
    return JSON.stringify({ ok: true });
  } catch (e) {
    return JSON.stringify({ ok: false, error: (e && e.message) ? e.message : String(e) });
  }
}

function insertCitation(data) {
  try {
    if (!data || typeof data !== 'object') return JSON.stringify({ ok: false, error: '삽입할 데이터가 없습니다.' });
    const saveResultStr = saveCitationOnly(data);
    const saveResult = typeof saveResultStr === 'string' ? JSON.parse(saveResultStr) : saveResultStr;
    if (saveResult && !saveResult.ok) return saveResultStr;

    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    const text = data.mode === 'manual' ? (data.fullText || '') : (data.shortText || '');
    const textToInsert = ' ' + text + ' ';

    let at = 'end';
    try {
      const cursor = doc.getCursor();
      if (cursor) {
        cursor.insertText(textToInsert);
        at = 'cursor';
      } else {
        const selection = doc.getSelection();
        if (selection && selection.getRangeElements().length > 0) {
          try {
            const rel = selection.getRangeElements()[0];
            const el = rel.getElement();
            const parent = el.getParent();
            if (parent && parent.getType() === DocumentApp.ElementType.BODY) {
              const idx = parent.getChildIndex(el);
              body.insertParagraph(idx + 1, textToInsert);
              at = 'cursor';
            } else {
              body.appendParagraph(textToInsert);
            }
          } catch (e) {
            body.appendParagraph(textToInsert);
          }
        } else {
          body.appendParagraph(textToInsert);
        }
      }
    } catch (docErr) {
      body.appendParagraph(textToInsert);
    }
    return JSON.stringify({ ok: true, at: at });
  } catch (e) {
    return JSON.stringify({ ok: false, error: (e && e.message) ? e.message : String(e) });
  }
}

function generateBibliography() {
  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();

    // 0. 이미 '참고문헌' 섹션이 있는지 검사 (제목 스타일 또는 문서 마지막 섹션인 경우만 오탐 방지)
    const numChildren = body.getNumChildren();
    var existingBibliographyError = "이미 '참고문헌' 섹션이 존재합니다. 기존 참고문헌 섹션을 지우거나 '참고문헌' 제목을 잠시 변경한 후 다시 시도해 주세요.";
    for (var i = 0; i < numChildren; i++) {
      var child = body.getChild(i);
      if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
      var p = child.asParagraph();
      if (p.getText().trim() !== '참고문헌') continue;
      var heading = p.getHeading();
      var isHeading = (heading !== DocumentApp.ParagraphHeading.NORMAL);
      var lastSectionThreshold = Math.max(5, Math.floor(numChildren * 0.1));
      var isInLastSection = (i >= numChildren - lastSectionThreshold);
      if (isHeading || isInLastSection) {
        return JSON.stringify({ ok: false, count: 0, message: existingBibliographyError, error: existingBibliographyError });
      }
    }

    // 1. 기존 참고문헌 구역 삭제 (위 조건에 해당하지 않는 경우만 여기 도달)
    var searchIndex = -1;
    for (var i = 0; i < numChildren; i++) {
      var child = body.getChild(i);
      if (child.getType() === DocumentApp.ElementType.PARAGRAPH && child.asParagraph().getText().trim() === '참고문헌') {
        searchIndex = i;
        break;
      }
    }
    if (searchIndex !== -1) {
      if (searchIndex > 0 && body.getChild(searchIndex - 1).getType() === DocumentApp.ElementType.PAGE_BREAK) searchIndex--;
      for (var j = body.getNumChildren() - 1; j >= searchIndex; j--) body.removeChild(body.getChild(j));
    }

    // 2. 데이터 불러오기 및 정렬
    const props = PropertiesService.getDocumentProperties();
    const citations = props.getProperty('CITATION_LIST');
    if (!citations) return JSON.stringify({ ok: false, count: 0, message: '저장된 참고문헌이 없습니다.' });
    let citationList = [];
    try {
      citationList = JSON.parse(citations);
    } catch (e) {
      return JSON.stringify({ ok: false, count: 0, message: '저장 데이터 오류.' });
    }
    if (!Array.isArray(citationList)) citationList = [];
    const uniqueMap = new Map();
    citationList.forEach(function(item) {
      const key = item.mode === 'manual'
        ? (item.author || '') + '|' + (item.secondVal || '') + '|' + (item.title || '')
        : (item.fullText || '');
      uniqueMap.set(key, item);
    });
    const uniqueCitations = Array.from(uniqueMap.values());
    uniqueCitations.sort(function(a, b) { return (a.author || a.fullText || '').localeCompare(b.author || b.fullText || ''); });

    // 3. 목록 작성 (서식 엔진 적용)
    var cfg = getEffectiveConfig();
    var headerAlign = cfg.headerAlignment === 'LEFT' ? DocumentApp.HorizontalAlignment.LEFT
      : cfg.headerAlignment === 'RIGHT' ? DocumentApp.HorizontalAlignment.RIGHT
      : DocumentApp.HorizontalAlignment.CENTER;
    var hangIndent = parseInt(cfg.hangingIndent, 10) || 20;
    var firstLineIndent = parseInt(cfg.indentFirstLine, 10) || 0;
    var lineSp = parseInt(cfg.lineSpacing, 10) || 12;

    body.appendPageBreak();
    var header = body.appendParagraph('참고문헌');
    header.setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(headerAlign);
    if (cfg.headerFontSize) header.editAsText().setFontSize(0, header.getText().length - 1, cfg.headerFontSize);

    var added = 0;
    uniqueCitations.forEach(function(item) {
      var result = formatCitationEntry(item, cfg);
      if (!result || !result.text || !result.text.trim()) return;
      var p = body.appendParagraph(result.text.trim());
      p.setIndentStart(hangIndent).setIndentFirstLine(firstLineIndent).setSpacingBefore(lineSp);
      if (result.ranges && result.ranges.length > 0) {
        var textObj = p.editAsText();
        var len = result.text.length;
        for (var r = 0; r < result.ranges.length; r++) {
          var range = result.ranges[r];
          var endIncl = range.end - 1;
          if (range.format === 'italics' && range.start >= 0 && endIncl < len && endIncl >= range.start) {
            textObj.setItalic(range.start, endIncl, true);
          }
        }
      }
      added++;
    });
    return JSON.stringify({ ok: true, count: added });
  } catch (e) {
    return JSON.stringify({ ok: false, count: 0, message: (e && e.message) ? e.message : String(e) });
  }
}

// ========== 더미 인용 감지 및 매핑 ==========

var SCORE_THRESHOLD = 3;

var DUMMY_SCAN_BLACKLIST = [
  '그림', '표', 'table', 'fig', 'figure', 'p.', 'pp.', 'page', 'pages',
  '참고', '주석', 'footnote', 'note', '각주', '출처', 'source', '이미지', 'image',
  '부록', 'appendix', '수식', 'equation', '식', '참조'
];

function isBlacklisted(innerContent) {
  if (!innerContent) return true;
  var content = innerContent.trim();
  for (var i = 0; i < DUMMY_SCAN_BLACKLIST.length; i++) {
    var term = DUMMY_SCAN_BLACKLIST[i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    if (new RegExp(term, 'i').test(content)) return true;
  }
  return false;
}

/**
 * 괄호 () 또는 대괄호 [] 안의 모든 짧은 텍스트(1~30자)를 잠재적 인용 후보로 감지
 * 한 글자 성(예: (오))도 감지. 그림, 표, p., pp. 등 인용이 아닌 괄호는 블랙리스트로 제외
 */
function scanDummyCitations() {
  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    const fullText = body.getText();

    const props = PropertiesService.getDocumentProperties();
    const citationsStr = props.getProperty('CITATION_LIST');
    let citationList = [];
    if (citationsStr) {
      try { citationList = JSON.parse(citationsStr); } catch (e) {}
    }
    if (!Array.isArray(citationList)) citationList = [];

    var re = /[\(\[][^()\[\]]{1,30}[\)\]]/g;
    var found = {};
    var dummyList = [];
    var m;

    while ((m = re.exec(fullText)) !== null) {
      var fullMatch = m[0];
      if (found[fullMatch]) continue;
      found[fullMatch] = true;

      var innerContent = fullMatch.slice(1, -1).trim();
      if (isBlacklisted(innerContent)) continue;

      var matchResult = fuzzyMatchCitation(citationList, innerContent, fullMatch);
      var isUnverified = matchResult.score < SCORE_THRESHOLD || matchResult.index < 0;
      dummyList.push({
        text: fullMatch,
        recommendedIndex: matchResult.index,
        score: matchResult.score,
        isUnverified: isUnverified
      });
    }

    var citationOptions = citationList.map(function(item, idx) {
      var label = '';
      if (item.mode === 'manual') {
        var author = (item.authorRaw || item.author || '').trim();
        var year = (item.secondVal || '').trim();
        var title = (item.title || '').trim();
        label = '[' + author + '(' + year + ')] ' + (title ? title.substring(0, 50) + (title.length > 50 ? '…' : '') : '(제목 없음)');
      } else {
        var short = (item.shortText || '').trim();
        var full = (item.fullText || '').trim();
        label = short ? '[' + short + '] ' + (full ? full.substring(0, 55) + (full.length > 55 ? '…' : '') : '') : (full ? full.substring(0, 80) + (full.length > 80 ? '…' : '') : '(내용 없음)');
      }
      return { index: idx, label: label };
    });

    return JSON.stringify({ ok: true, dummies: dummyList, citationOptions: citationOptions, threshold: SCORE_THRESHOLD });
  } catch (e) {
    return JSON.stringify({ ok: false, error: (e && e.message) ? e.message : String(e) });
  }
}

/**
 * 가중치 기반 스코어링: 저자 +5, 연도 +5, 제목키워드 단어당 +2
 * @return {{ index: number, score: number }}
 */
function fuzzyMatchCitation(citationList, innerContent, fullMatch) {
  if (citationList.length === 0) return { index: -1, score: 0 };
  var content = (innerContent || '').toLowerCase();

  var bestIdx = -1;
  var bestScore = 0;

  for (var i = 0; i < citationList.length; i++) {
    var c = citationList[i];
    var score = 0;

    var surname = extractSurname(c.author);
    if (surname && content.indexOf(surname.toLowerCase()) >= 0) score += 5;

    var yr = String(c.secondVal || '');
    if (yr.length >= 2) {
      var yrFull = yr;
      var yrShort = yr.length === 4 ? yr.slice(2) : yr;
      if (content.indexOf(yrFull) >= 0 || content.indexOf(yrShort) >= 0) score += 5;
    }

    var titleWords = (c.title || '').match(/[A-Za-z가-힣\uac00-\ud7a3]{3,}/g) || [];
    for (var w = 0; w < titleWords.length; w++) {
      if (content.indexOf(titleWords[w].toLowerCase()) >= 0) score += 2;
    }

    if (c.fullText === fullMatch || (c.shortText && c.shortText === fullMatch)) score += 10;

    if (score > bestScore) {
      bestScore = score;
      bestIdx = i;
    }
  }
  return { index: bestScore > 0 ? bestIdx : -1, score: bestScore };
}

function extractSurname(authorStr) {
  if (!authorStr || typeof authorStr !== 'string') return '';
  var s = authorStr.trim();
  var commaIdx = s.indexOf(',');
  if (commaIdx >= 0) return s.substring(0, commaIdx).trim();
  var semiIdx = s.indexOf(';');
  if (semiIdx >= 0) return s.substring(0, semiIdx).trim();
  var spaceIdx = s.indexOf(' ');
  if (spaceIdx >= 0) return s.substring(0, spaceIdx).trim();
  if (s.length >= 2 && /[\uac00-\ud7a3\u4e00-\u9fff]/.test(s[0])) return s[0];
  return s;
}

/**
 * 사용자 매핑에 따라 문서 내 더미 인용을 정식 인용구로 일괄 치환.
 * citationIndex === -1(대체하지 않음/Skip)인 항목은 치환하지 않음.
 * 치환된 인용구는 폰트 색상을 빨간색(#FF0000)으로 설정하여 시각적 피드백 제공.
 */
function applyMapping(mappings) {
  try {
    if (!mappings || !Array.isArray(mappings) || mappings.length === 0) {
      return JSON.stringify({ ok: false, error: '매핑 데이터가 없습니다.' });
    }

    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    const props = PropertiesService.getDocumentProperties();
    const citationsStr = props.getProperty('CITATION_LIST');
    let citationList = [];
    if (citationsStr) {
      try { citationList = JSON.parse(citationsStr); } catch (e) {}
    }
    if (!Array.isArray(citationList)) citationList = [];

    var replaced = 0;
    for (var i = 0; i < mappings.length; i++) {
      var m = mappings[i];
      var citationIndex = parseInt(m.citationIndex, 10);
      if (citationIndex === -1 || isNaN(citationIndex) || citationIndex < 0 || citationIndex >= citationList.length) continue;

      var citation = citationList[citationIndex];
      var replacementText = '';
      if (citation.mode === 'manual') {
        replacementText = getInTextCitationString(
          citation.authorRaw || citation.author,
          citation.secondVal || '',
          citation.style || 'APA'
        );
      } else {
        replacementText = citation.shortText || citation.fullText || '';
      }
      if (!replacementText) continue;

      var searchText = String(m.dummyText || '').trim();
      if (!searchText) continue;

      var escaped = escapeRegexForReplace(searchText);
      var safeReplacement = escapeReplacementText(replacementText);
      body.replaceText(escaped, safeReplacement);
      var replEscaped = escapeRegexForReplace(safeReplacement);
      var range = body.findText(replEscaped);
      while (range) {
        var el = range.getElement();
        var start = range.getStartOffset();
        var end = range.getEndOffsetInclusive();
        try {
          el.asText().setForegroundColor(start, end, '#FF0000');
        } catch (colorErr) {}
        replaced++;
        range = body.findText(replEscaped, range);
      }
    }

    return JSON.stringify({ ok: true, replaced: replaced });
  } catch (e) {
    return JSON.stringify({ ok: false, error: (e && e.message) ? e.message : String(e) });
  }
}

/** 정규식 특수문자 이스케이프 (replaceText 검색 패턴용) */
function escapeRegexForReplace(str) {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** 치환 문자열 내 $, \\ 이스케이프 (replaceText 치환문 오류 방지) */
function escapeReplacementText(str) {
  return String(str).replace(/\\/g, '\\\\').replace(/\$/g, '$$');
}