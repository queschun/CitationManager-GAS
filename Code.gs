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
  var L = getLanguagePack();
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
    var detail = (e && e.message) ? e.message : String(e);
    return JSON.stringify({ ok: false, error: formatMessage(L.errors.operationFailed, { detail: detail }) });
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
  var L = getLanguagePack();
  try {
    if (!fullText || typeof fullText !== 'string') {
      return JSON.stringify({ ok: false, error: L.errors.noTextToParse });
    }
    var text = fullText.trim();
    if (!text) {
      return JSON.stringify({ ok: false, error: L.errors.noTextToParse });
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
    return JSON.stringify({ ok: false, error: formatMessage(L.errors.parseError, { detail: errMsg }) });
  }
}

/**
 * 참고문헌용 저자 리스트 생성: 저자 생략 금지, 마지막 저자 앞 & 필수, 클렌징 적용.
 * CITATION_LIST에 저자가 긴 문자열 하나로 저장돼 있어도 쉼표로 분리 후 재조립.
 * @param {string} authorStr - 저자 문자열 (세미콜론 또는 쉼표 구분)
 * @return {string} "저자1, 저자2, & 저자3" 형태 (전원 포함, 생략 없음)
 */
function formatAuthorListForBibliography(authorStr) {
  if (authorStr == null || typeof authorStr !== 'string') return 'Unknown';
  var raw = String(authorStr).trim();
  if (!raw) return 'Unknown';
  var parts;
  if (raw.indexOf(';') >= 0) {
    parts = raw.split(';');
  } else {
    parts = raw.split(',');
  }
  parts = parts.map(function(s) {
    return s.replace(/\s+/g, ' ').replace(/^[\s,]+|[\s,]+$/g, '').trim();
  }).filter(function(s) { return s.length > 0; });
  if (parts.length === 0) return 'Unknown';
  if (parts.length === 1) return parts[0];
  return parts.slice(0, -1).join(', ') + ', & ' + parts[parts.length - 1];
}

/**
 * 저자 문자열을 STYLE_CONFIG에 따라 포맷.
 * 참고문헌(isBibliography)일 때: formatAuthorListForBibliography로 전원 포함·마지막 앞 & 고정·생략 금지.
 * 본문 인용일 때: 기존 etAlThreshold/et al.·외 로직 유지.
 */
function formatAuthorsWithConfig(authorStr, cfg, isKo, isBibliography) {
  if (!authorStr) return 'Unknown';
  if (isBibliography) {
    return formatAuthorListForBibliography(authorStr);
  }
  var parts = authorStr.split(';').map(function(s) { return s.trim(); }).filter(function(s) { return s; });
  if (parts.length === 0) return 'Unknown';
  var sep = (cfg && cfg.authorSeparator !== undefined) ? cfg.authorSeparator : ', ';
  var etAl = isKo ? (cfg && cfg.etAlKo) ? cfg.etAlKo : '외' : (cfg && cfg.etAlEn) ? cfg.etAlEn : 'et al.';
  var th = (cfg && parseInt(cfg.etAlThreshold, 10)) ? parseInt(cfg.etAlThreshold, 10) : 3;
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
  var L = getLanguagePack();
  DocumentApp.getUi().createMenu(L.ui.menu.title)
      .addItem(L.ui.menu.openSidebar, 'showSidebar')
      .addSeparator()
      .addItem(L.ui.menu.clearData, 'clearCitations')
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
  var L = getLanguagePack();
  try {
    if (!data || typeof data !== 'object') return JSON.stringify({ ok: false, error: L.errors.noDataToSave });
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
    if (!isDuplicate(data, citationList)) citationList.push(data);
    props.setProperty('CITATION_LIST', JSON.stringify(citationList));
    return JSON.stringify({ ok: true });
  } catch (e) {
    var detail = (e && e.message) ? e.message : String(e);
    return JSON.stringify({ ok: false, error: formatMessage(L.errors.operationFailed, { detail: detail }) });
  }
}

/**
 * 저장된 문헌 목록 반환. 사이드바 문헌 리스트(체크박스) 표시용.
 * @return {string} JSON. { ok: true, list: Array } 또는 { ok: false, error: string }
 */
function getCitationList() {
  var L = getLanguagePack();
  try {
    var props = PropertiesService.getDocumentProperties();
    var citations = props.getProperty('CITATION_LIST');
    var list = [];
    if (citations) {
      try {
        list = JSON.parse(citations);
      } catch (e) {
        list = [];
      }
    }
    if (!Array.isArray(list)) list = [];
    return JSON.stringify({ ok: true, list: list });
  } catch (e) {
    var detail = (e && e.message) ? e.message : String(e);
    return JSON.stringify({ ok: false, error: formatMessage(L.errors.operationFailed, { detail: detail }) });
  }
}

/**
 * 새 문헌이 기존 목록에 이미 있는지 판별. 저자·연도·제목이 모두 일치하면 중복.
 * @param {Object} newCitation - 추가하려는 문헌 객체
 * @param {Array.<Object>} existingList - 기존 CITATION_LIST 배열
 * @return {boolean} true면 중복(추가하지 말 것)
 */
function isDuplicate(newCitation, existingList) {
  if (!newCitation || typeof newCitation !== 'object') return false;
  if (!existingList || !Array.isArray(existingList)) return false;
  var key = getCitationDedupeKey(newCitation);
  if (!key) return false;
  var keyNorm = key.toLowerCase().replace(/\s+/g, ' ');
  for (var i = 0; i < existingList.length; i++) {
    var existing = existingList[i];
    var existingKey = getCitationDedupeKey(existing);
    if (!existingKey) continue;
    if (existingKey.toLowerCase().replace(/\s+/g, ' ') === keyNorm) return true;
  }
  return false;
}

/**
 * 문헌 객체에서 '같은 문헌' 판별용 정규화 키 반환. 지능형 중복 제거용.
 * manual: 저자|연도|제목, paste: shortText 또는 fullText(trim).
 * @param {Object} item - 문헌 객체
 * @return {string} 정규화된 키 (빈 문자열이면 무시 대상)
 */
function getCitationDedupeKey(item) {
  if (!item || typeof item !== 'object') return '';
  if (item.mode === 'manual') {
    var author = String(item.authorRaw || item.author || '').trim();
    var year = String(item.secondVal || '').trim();
    var title = String(item.title || '').trim();
    return author + '|' + year + '|' + title;
  }
  var shortText = String(item.shortText || '').trim();
  if (shortText) return shortText;
  var fullText = String(item.fullText || '').trim();
  return fullText;
}

/**
 * citationArray에서 동일 문헌 중복을 제거. 첫 번째 출현만 유지(순서 유지).
 * @param {Array.<Object>} citationArray - 문헌 객체 배열
 * @return {Array.<Object>} 중복 제거된 배열
 */
function deduplicateCitations(citationArray) {
  if (!citationArray || !Array.isArray(citationArray)) return [];
  var seen = {};
  var result = [];
  for (var i = 0; i < citationArray.length; i++) {
    var item = citationArray[i];
    var key = getCitationDedupeKey(item);
    if (!key) continue;
    var keyNorm = key.toLowerCase().replace(/\s+/g, ' ');
    if (seen[keyNorm]) continue;
    seen[keyNorm] = true;
    result.push(item);
  }
  return result;
}

/**
 * 인용구 알맹이 추출: 괄호 ((, ), [, ]) 모두 제거, 양 끝 공백 제거 후 순수 "저자, 연도"만 반환.
 * @param {string|Object} citationData - 인용 문자열 또는 문헌 객체
 * @return {string} "저자, 연도" 형태 (괄호 없음)
 */
function getCleanText(citationData) {
  if (citationData == null) return '';
  var text = '';
  if (typeof citationData === 'string') {
    text = String(citationData);
  } else if (typeof citationData === 'object') {
    text = citationData.shortCitation || getShortCitationPart(citationData) || '';
  }
  return String(text).replace(/[()\[\]]/g, '').trim();
}

/**
 * Universal Sanitizer: ((...); (...)) 형태를 (저자1; 저자2) 형태로 정규화.
 * - 문자열 내 모든 (, ), [, ] 제거 후 알맹이만 유지
 * - 세미콜론(;)으로 분리하여 배열화
 * - 동일 저자/연도 중복 제거 (정규화 키로 필터)
 * - 알맹이들을 "; "로 합치고 앞뒤에 ( ) 한 번만
 * @param {string} str - 임의의 인용 문자열
 * @return {string} "(저자1, 연도; 저자2, 연도)" 형태
 */
function cleanCitationFormat(str) {
  if (str == null || typeof str !== 'string') return '';
  var raw = String(str).replace(/[()\[\]]/g, '').trim();
  if (!raw) return '';
  var parts = raw.split(';').map(function(s) { return s.trim(); }).filter(function(s) { return s.length > 0; });
  var seen = {};
  var unique = [];
  for (var i = 0; i < parts.length; i++) {
    var p = parts[i];
    var key = p.toLowerCase().replace(/\s+/g, ' ');
    if (seen[key]) continue;
    seen[key] = true;
    unique.push(p);
  }
  if (unique.length === 0) return '';
  return '(' + unique.join('; ') + ')';
}

/**
 * 문헌 한 건에서 Short Citation(저자, 연도) 부분만 추출. 괄호 없이 "저자, 연도" 문자열 반환.
 * @param {Object} item - 문헌 객체 (mode: 'manual' | 'paste')
 * @return {string} "저자, 연도" 형식
 */
function getShortCitationPart(item) {
  if (!item) return '';
  if (item.mode === 'manual') {
    var authorRaw = (item.authorRaw || item.author || '').trim();
    var year = (item.secondVal || '').trim();
    var style = (item.style || 'APA').toUpperCase();
    var isKo = isKoreanCitation(authorRaw);
    var author = formatAuthorsForInText(authorRaw, isKo);
    return style === 'APA' ? author + ', ' + year : author + ' ' + year;
  }
  var shortText = (item.shortText || '').trim();
  if (shortText) {
    var stripped = shortText.replace(/^\s*\(?\s*|\s*\)?\s*$/g, '');
    return stripped || shortText;
  }
  var fullText = (item.fullText || '').trim();
  if (fullText) {
    var suggested = parseAndSuggestCitation(fullText);
    try {
      var parsed = typeof suggested === 'string' ? JSON.parse(suggested) : suggested;
      if (parsed && parsed.ok && parsed.citationText) {
        var ct = parsed.citationText.replace(/^\s*\(?\s*|\s*\)?\s*$/g, '');
        return ct || fullText.substring(0, 50);
      }
    } catch (e) {}
    return fullText.substring(0, 50);
  }
  return '';
}

/**
 * 다중 인용 및 DB 중복 방지 통합 함수
 */
function insertMultiCitation(citationArray) {
  if (!citationArray || citationArray.length === 0) return;

  // 1. 중복 데이터 제거 (입력 배열 내 중복)
  const uniqueItems = deduplicateCitations(citationArray);

  // 2. 알맹이 추출 (getCleanText: 괄호 제거·양끝 공백 제거 → "저자, 연도"만)
  var cleanedParts = uniqueItems.map(function(item) {
    return getCleanText(item);
  }).filter(function(s) { return s.length > 0; });

  if (cleanedParts.length === 0) return;

  // 3. 단일 괄호로 재조립 후 Universal Sanitizer 적용
  var assembled = "(" + cleanedParts.join("; ") + ")";
  var finalString = cleanCitationFormat(assembled);
  if (!finalString) return;

  // 4. 문서 삽입 (괄호 포함 전체 빨간색 #ff0000 강제)
  var doc = DocumentApp.getActiveDocument();
  var cursor = doc.getCursor();

  if (cursor) {
    var startOffset = cursor.getOffset();
    cursor.insertText(finalString);
    cursor.getElement().asText().setForegroundColor(startOffset, startOffset + finalString.length - 1, "#ff0000");
  } else {
    var para = doc.getBody().appendParagraph(finalString);
    para.editAsText().setForegroundColor(0, finalString.length - 1, "#ff0000");
  }

  // 5. DB 중복 체크 후 저장
  saveToCitationListWithDedupe(uniqueItems);

  return JSON.stringify({ ok: true, count: cleanedParts.length });
}

/**
 * DB 저장 시 중복 체크 로직: some()으로 저자·연도·제목이 모두 일치하면 push 생략.
 */
function saveToCitationListWithDedupe(newItems) {
  var props = PropertiesService.getDocumentProperties();
  var currentList = [];
  try {
    currentList = JSON.parse(props.getProperty('CITATION_LIST') || '[]');
  } catch (e) {
    currentList = [];
  }
  if (!Array.isArray(currentList)) currentList = [];

  for (var i = 0; i < newItems.length; i++) {
    var newItem = newItems[i];
    var na = (newItem.author || newItem.authorRaw || '').trim();
    var ny = (newItem.year != null ? newItem.year : newItem.secondVal || '').toString().trim();
    var nt = (newItem.title || '').trim();

    var alreadyExists = currentList.some(function(existItem) {
      var ea = (existItem.author || existItem.authorRaw || '').trim();
      var ey = (existItem.year != null ? existItem.year : existItem.secondVal || '').toString().trim();
      var et = (existItem.title || '').trim();
      return ea === na && ey === ny && et === nt;
    });

    if (!alreadyExists) {
      currentList.push(newItem);
    }
  }

  props.setProperty('CITATION_LIST', JSON.stringify(currentList));
}

function insertCitation(data) {
  var L = getLanguagePack();
  try {
    if (!data || typeof data !== 'object') return JSON.stringify({ ok: false, error: L.errors.noDataToInsert });
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
    var detail = (e && e.message) ? e.message : String(e);
    return JSON.stringify({ ok: false, error: formatMessage(L.errors.operationFailed, { detail: detail }) });
  }
}

function generateBibliography() {
  var L = getLanguagePack();
  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();

    // 0. 이미 '참고문헌' 섹션이 있는지 검사 (제목 스타일 또는 문서 마지막 섹션인 경우만 오탐 방지)
    const numChildren = body.getNumChildren();
    var existingBibliographyError = L.errors.bibliographyExists;
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
    if (!citations) return JSON.stringify({ ok: false, count: 0, message: L.errors.noCitationsStored, error: L.errors.noCitationsStored });
    let citationList = [];
    try {
      citationList = JSON.parse(citations);
    } catch (e) {
      return JSON.stringify({ ok: false, count: 0, message: L.errors.citationDataError, error: L.errors.citationDataError });
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
    var detail = (e && e.message) ? e.message : String(e);
    return JSON.stringify({ ok: false, count: 0, message: formatMessage(L.errors.operationFailed, { detail: detail }), error: formatMessage(L.errors.operationFailed, { detail: detail }) });
  }
}

/**
 * 참고문헌 섹션을 정렬·서식 적용한다.
 * 바구니 통합: koBucket.concat(enBucket, otherBucket)으로 finalSortedList 생성. 빈 텍스트 방어 후 setText/appendParagraph 호출.
 * 남는 단락은 마지막 단락(isLastParagraph)이면 setText(" ")만, 아니면 삭제. 모든 문헌에 내어쓰기 20pt·줄간격 1.5 적용.
 * @return {string} JSON. { ok: true, count: number, message: string } 또는 { ok: false, error: string }
 */
function sortAndFormatBibliography() {
  var L = getLanguagePack();
  var HANG_INDENT_PT = 20;
  var LINE_SPACING = 1.5;
  var SAFE_SPACE = ' ';

  function safeText(val) {
    if (val == null || val === undefined) return SAFE_SPACE;
    var s = String(val).trim();
    return s !== '' ? s : SAFE_SPACE;
  }

  try {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var numChildren = body.getNumChildren();

    var refIndex = -1;
    for (var i = 0; i < numChildren; i++) {
      var child = body.getChild(i);
      if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
        var titleText = child.asParagraph().getText();
        if (titleText != null && String(titleText).trim() === '참고문헌') {
          refIndex = i;
          break;
        }
      }
    }
    if (refIndex < 0) {
      return JSON.stringify({ ok: false, error: L.errors.bibliographyExists || '참고문헌 섹션을 찾을 수 없습니다.' });
    }

    // 수집: 제목 아래 문헌만, trim() 후 내용 있는 것만
    var raw = [];
    for (var j = refIndex + 1; j < numChildren; j++) {
      var el = body.getChild(j);
      if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
        var text = el.asParagraph().getText();
        if (text != null) text = String(text).trim();
        if (text && text !== '') raw.push(text);
      }
    }

    // 바구니 분류: 첫 글자 기준 ko / en / other. 각 바구니에서 trim 후 내용 있는 것만 유지
    var koBucket = [];
    var enBucket = [];
    var otherBucket = [];
    for (var r = 0; r < raw.length; r++) {
      var s = raw[r];
      if (s == null || String(s).trim() === '') continue;
      s = String(s).trim();
      if (s === '') continue;
      var first = (s.length > 0) ? s.charAt(0) : '';
      if (/[\uac00-\ud7a3]/.test(first)) {
        koBucket.push(s);
      } else if (/[A-Za-z]/.test(first)) {
        enBucket.push(s);
      } else {
        otherBucket.push(s);
      }
    }
    koBucket = koBucket.filter(function(item) { return item != null && String(item).trim() !== ''; });
    enBucket = enBucket.filter(function(item) { return item != null && String(item).trim() !== ''; });
    otherBucket = otherBucket.filter(function(item) { return item != null && String(item).trim() !== ''; });

    koBucket.sort(function(a, b) { return a.localeCompare(b, 'ko'); });
    enBucket.sort(function(a, b) { return a.localeCompare(b, 'en'); });
    otherBucket.sort(function(a, b) { return a.localeCompare(b, 'zh-CN'); });

    // 바구니 통합 (The Missing Link): 반드시 세 바구니를 하나로 합침. 영문 누락 방지.
    var finalSortedList = koBucket.concat(enBucket, otherBucket);

    // 제목 아래 단락 개수(문단만)
    var paraIndices = [];
    for (var j = refIndex + 1; j < body.getNumChildren(); j++) {
      if (body.getChild(j).getType() === DocumentApp.ElementType.PARAGRAPH) {
        paraIndices.push(j);
      }
    }

    // 단락이 부족하면 appendParagraph(" ")로 확보. 전달값은 절대 null/undefined/"" 아님.
    var listLen = finalSortedList.length > 0 ? finalSortedList.length : 1;
    while (paraIndices.length < listLen) {
      body.appendParagraph(SAFE_SPACE);
      paraIndices = [];
      for (var j = refIndex + 1; j < body.getNumChildren(); j++) {
        if (body.getChild(j).getType() === DocumentApp.ElementType.PARAGRAPH) paraIndices.push(j);
      }
    }

    // 안전한 덮어쓰기: 제목 아래 단락들을 순회하며 finalSortedList 값으로 채움. setText(null/undefined/"") 호출 금지.
    for (var n = 0; n < listLen && n < paraIndices.length; n++) {
      var idx = paraIndices[n];
      var child = body.getChild(idx);
      if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
      var val = (n < finalSortedList.length && finalSortedList[n] != null) ? finalSortedList[n] : '';
      var toSet = safeText(val);
      child.asParagraph().setText(toSet);
    }

    // 남는 단락 정리: finalSortedList보다 단락이 많으면 삭제. 문서의 진짜 마지막 단락(isLastParagraph)이면 삭제하지 말고 setText(" ")만
    paraIndices = [];
    for (var j = refIndex + 1; j < body.getNumChildren(); j++) {
      if (body.getChild(j).getType() === DocumentApp.ElementType.PARAGRAPH) paraIndices.push(j);
    }
    for (var i = paraIndices.length - 1; i >= listLen; i--) {
      var k = paraIndices[i];
      var c = body.getChild(k);
      var isLastParagraph = (k === body.getNumChildren() - 1);
      if (isLastParagraph) {
        if (c.getType() === DocumentApp.ElementType.PARAGRAPH) {
          c.asParagraph().setText(SAFE_SPACE);
        }
        break;
      }
      body.removeChild(c);
    }

    // 서식 강제 적용: 모든 참고문헌 단락에 setIndentationHanging(20)·setLineSpacing(1.5)
    function applyFormat(p) {
      if (p == null) return;
      p.setIndentFirstLine(-HANG_INDENT_PT);
      p.setIndentStart(HANG_INDENT_PT);
      p.setSpacingBefore(LINE_SPACING * 6);
      p.setSpacingAfter(0);
    }
    paraIndices = [];
    for (var j = refIndex + 1; j < body.getNumChildren(); j++) {
      if (body.getChild(j).getType() === DocumentApp.ElementType.PARAGRAPH) paraIndices.push(j);
    }
    for (var m = 0; m < listLen && m < paraIndices.length; m++) {
      var node = body.getChild(paraIndices[m]);
      if (node.getType() === DocumentApp.ElementType.PARAGRAPH) {
        applyFormat(node.asParagraph());
      }
    }

    var doneMessage = (L.messages && L.messages.sortKoFirstDone) ? L.messages.sortKoFirstDone : '국문 우선 정렬 및 서식 적용 완료';
    return JSON.stringify({ ok: true, count: finalSortedList.length, message: doneMessage });
  } catch (e) {
    var detail = (e && e.message) ? e.message : String(e);
    return JSON.stringify({ ok: false, error: formatMessage(L.errors.operationFailed, { detail: detail }) });
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
  var L = getLanguagePack();
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

      if (innerContent.indexOf(';') >= 0) {
        var parts = innerContent.split(';').map(function(s) { return s.trim(); }).filter(function(s) { return s.length > 0; });
        for (var p = 0; p < parts.length; p++) {
          var part = parts[p];
          if (isBlacklisted(part)) continue;
          var matchResult = fuzzyMatchCitation(citationList, part, fullMatch);
          var isUnverified = matchResult.score < SCORE_THRESHOLD || matchResult.index < 0;
          dummyList.push({
            text: part,
            recommendedIndex: matchResult.index,
            score: matchResult.score,
            isUnverified: isUnverified
          });
        }
      } else {
        var matchResult = fuzzyMatchCitation(citationList, innerContent, fullMatch);
        var isUnverified = matchResult.score < SCORE_THRESHOLD || matchResult.index < 0;
        dummyList.push({
          text: fullMatch,
          recommendedIndex: matchResult.index,
          score: matchResult.score,
          isUnverified: isUnverified
        });
      }
    }

    var citationOptions = citationList.map(function(item, idx) {
      var label = '';
      if (item.mode === 'manual') {
        var author = (item.authorRaw || item.author || '').trim();
        var year = (item.secondVal || '').trim();
        var title = (item.title || '').trim();
        label = '[' + author + '(' + year + ')] ' + (title ? title.substring(0, 50) + (title.length > 50 ? '…' : '') : L.errors.noTitle);
      } else {
        var short = (item.shortText || '').trim();
        var full = (item.fullText || '').trim();
        label = short ? '[' + short + '] ' + (full ? full.substring(0, 55) + (full.length > 55 ? '…' : '') : '') : (full ? full.substring(0, 80) + (full.length > 80 ? '…' : '') : L.errors.noContent);
      }
      return { index: idx, label: label };
    });

    return JSON.stringify({ ok: true, dummies: dummyList, citationOptions: citationOptions, threshold: SCORE_THRESHOLD });
  } catch (e) {
    var detail = (e && e.message) ? e.message : String(e);
    return JSON.stringify({ ok: false, error: formatMessage(L.errors.operationFailed, { detail: detail }) });
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
 * 교체 후 Universal Sanitizer(cleanCitationFormat)로 괄호 안 중복·이중괄호 정리,
 * 최종 문자열(괄호 포함) 전체를 빨간색(#ff0000)으로 강제.
 */
function applyMapping(mappingList) {
  var L = getLanguagePack();
  try {
    if (!mappingList || !Array.isArray(mappingList) || mappingList.length === 0) {
      return JSON.stringify({ ok: false, error: L.errors.noMappingData });
    }

    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var props = PropertiesService.getDocumentProperties();
    var citationList = [];
    try {
      var saved = props.getProperty('CITATION_LIST');
      if (saved) citationList = JSON.parse(saved);
    } catch (e) {}
    if (!Array.isArray(citationList)) citationList = [];

    var replaced = 0;
    for (var i = 0; i < mappingList.length; i++) {
      var m = mappingList[i];
      var dummyText = String(m.dummyText || '').trim();
      var realData = m.realData;
      if (!dummyText) continue;
      if (!realData || typeof realData !== 'object') continue;

      var replacementText = '';
      if (realData.mode === 'manual') {
        replacementText = getInTextCitationString(
          realData.authorRaw || realData.author,
          realData.secondVal || '',
          realData.style || 'APA'
        );
      } else {
        replacementText = (realData.shortText || realData.fullText || '').trim();
      }
      if (!replacementText) continue;

      var escapedSearch = escapeRegexForReplace(dummyText);
      var safeReplacement = escapeReplacementText(replacementText);
      body.replaceText(escapedSearch, safeReplacement);

      var replEscaped = escapeRegexForReplace(safeReplacement);
      var range = body.findText(replEscaped);
      while (range) {
        var el = range.getElement();
        var start = range.getStartOffset();
        var end = range.getEndOffsetInclusive();
        try {
          el.asText().setForegroundColor(start, end, '#ff0000');
          replaced++;
        } catch (colorErr) {}
        range = body.findText(replEscaped, range);
      }

      if (!citationList.some(function(exist) {
        var a = (exist.author || exist.authorRaw || '').trim();
        var y = (exist.year != null ? exist.year : exist.secondVal || '').toString().trim();
        var t = (exist.title || '').trim();
        var na = (realData.author || realData.authorRaw || '').trim();
        var ny = (realData.year != null ? realData.year : realData.secondVal || '').toString().trim();
        var nt = (realData.title || '').trim();
        return a === na && y === ny && t === nt;
      })) {
        citationList.push(realData);
      }
    }

    props.setProperty('CITATION_LIST', JSON.stringify(citationList));

    // 교체가 일어난 전체 단락·괄호 덩어리 재읽기: ((A); (B)) 정규식으로 찾아 내부 괄호 제거·중복 제거·바깥 괄호만 빨간색
    applyCleanCitationFormatToBody(body);

    return JSON.stringify({ ok: true, replaced: replaced });
  } catch (e) {
    var detail = (e && e.message) ? e.message : String(e);
    return JSON.stringify({ ok: false, error: formatMessage(L.errors.operationFailed, { detail: detail }) });
  }
}

/**
 * body 내 텍스트 요소를 순회하며 인용 괄호 덩어리(단일 또는 ((A); (B)) 중첩)를 찾아
 * cleanCitationFormat 적용 후 덮어쓰고, 괄호부터 세미콜론까지 전부 빨간색(#ff0000) 강제.
 */
function applyCleanCitationFormatToBody(body) {
  var RED = '#ff0000';
  var numChildren = body.getNumChildren();
  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    var type = child.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      applyCleanCitationFormatToText(child.asParagraph(), RED);
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      applyCleanCitationFormatToText(child.asListItem(), RED);
    } else if (type === DocumentApp.ElementType.TABLE) {
      var table = child.asTable();
      for (var r = 0; r < table.getNumRows(); r++) {
        var row = table.getRow(r);
        for (var c = 0; c < row.getNumCells(); c++) {
          var cell = row.getCell(c);
          for (var cc = 0; cc < cell.getNumChildren(); cc++) {
            var cellChild = cell.getChild(cc);
            if (cellChild.getType() === DocumentApp.ElementType.PARAGRAPH) {
              applyCleanCitationFormatToText(cellChild.asParagraph(), RED);
            }
          }
        }
      }
    }
  }
}

/**
 * 단일 텍스트 요소 내 ( ... ) 또는 ((A); (B)) 형태 블록을 정규식으로 찾아
 * 내부 괄호 제거·동일 저자+연도 중복 제거 후 바깥 괄호만 씌워 덮어쓰고, 전체 빨간색 적용.
 */
function applyCleanCitationFormatToText(textEl, color) {
  if (!textEl || textEl.editAsText == null) return;
  var rich = textEl.editAsText();
  var text = rich.getText();
  if (!text) return;
  var matches = [];
  var re = /\((?:(?:[^()]+)|(?:\([^()]*\)))*\)/g;
  var m;
  while ((m = re.exec(text)) !== null) {
    matches.push({ start: m.index, end: m.index + m[0].length - 1, original: m[0] });
  }
  for (var i = matches.length - 1; i >= 0; i--) {
    var cleaned = cleanCitationFormat(matches[i].original);
    if (!cleaned || cleaned === matches[i].original) continue;
    var start = matches[i].start;
    var end = matches[i].end;
    try {
      rich.deleteText(start, end);
      rich.insertText(start, cleaned);
      rich.setForegroundColor(start, start + cleaned.length - 1, color || '#ff0000');
    } catch (err) {}
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

/**
 * 문서 전체에서 글자 색상이 빨간색(rgb(255,0,0) / #ff0000)인 텍스트를 검정색으로 바꾼다.
 * 변경된 연속 구간(런) 개수를 반환한다.
 * @return {number} 변경된 구간 개수
 */
function finalizeCitations() {
  var RED = '#ff0000';
  var BLACK = '#000000';
  var changedCount = 0;
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  function processText(textEl) {
    if (!textEl || textEl.editAsText === undefined) return;
    var text = textEl.editAsText();
    var len = text.getText().length;
    if (len === 0) return;
    var start = 0;
    while (start < len) {
      var color = (text.getForegroundColor(start) || '').toString().toLowerCase();
      if (color === RED || color === '#ff0000') {
        var end = start;
        while (end < len) {
          var c = (text.getForegroundColor(end) || '').toString().toLowerCase();
          if (c !== RED && c !== '#ff0000') break;
          end++;
        }
        if (end > start) {
          text.setForegroundColor(start, end - 1, BLACK);
          changedCount++;
        }
        start = end;
      } else {
        start++;
      }
    }
  }

  var numChildren = body.getNumChildren();
  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    var type = child.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      processText(child.asParagraph());
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      processText(child.asListItem());
    } else if (type === DocumentApp.ElementType.TABLE) {
      var table = child.asTable();
      var rows = table.getNumRows();
      for (var r = 0; r < rows; r++) {
        var row = table.getRow(r);
        var cells = row.getNumCells();
        for (var c = 0; c < cells; c++) {
          var cell = row.getCell(c);
          var cellNumChildren = cell.getNumChildren();
          for (var cc = 0; cc < cellNumChildren; cc++) {
            var cellChild = cell.getChild(cc);
            if (cellChild.getType() === DocumentApp.ElementType.PARAGRAPH) {
              processText(cellChild.asParagraph());
            }
          }
        }
      }
    }
  }

  return changedCount;
}