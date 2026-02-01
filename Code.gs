/** @OnlyCurrentDoc */

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

    // 1. 기존 참고문헌 구역 삭제
    const numChildren = body.getNumChildren();
    let searchIndex = -1;
    for (let i = 0; i < numChildren; i++) {
      const child = body.getChild(i);
      if (child.getType() === DocumentApp.ElementType.PARAGRAPH && child.asParagraph().getText().trim() === '참고문헌') {
        searchIndex = i;
        break;
      }
    }
    if (searchIndex !== -1) {
      if (searchIndex > 0 && body.getChild(searchIndex - 1).getType() === DocumentApp.ElementType.PAGE_BREAK) searchIndex--;
      for (let j = body.getNumChildren() - 1; j >= searchIndex; j--) body.removeChild(body.getChild(j));
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

    // 3. 목록 작성 (APA / MLA 양식별)
    body.appendPageBreak();
    const header = body.appendParagraph('참고문헌');
    header.setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    let added = 0;
    uniqueCitations.forEach(function(item) {
      let entry = '';
      if (item && item.mode === 'manual') {
        const style = item.style || 'APA';
        const author = item.author || '';
        const title = item.title || '';
        const year = item.secondVal || '';
        const publisher = item.publisher || '';
        if (style === 'MLA') {
          entry = author + '. "' + title + '." ' + (publisher ? publisher + ', ' : '') + year + '.';
        } else {
          entry = author + ' (' + year + '). ' + title + '. ' + publisher;
        }
      } else {
        entry = (item && item.fullText) ? String(item.fullText) : '';
      }
      if (!entry.trim()) return;
      const p = body.appendParagraph(entry.trim());
      p.setIndentStart(20).setIndentFirstLine(0).setSpacingBefore(12);
      added++;
    });
    return JSON.stringify({ ok: true, count: added });
  } catch (e) {
    return JSON.stringify({ ok: false, count: 0, message: (e && e.message) ? e.message : String(e) });
  }
}