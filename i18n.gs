/** @OnlyCurrentDoc */
/**
 * i18n: 한국어(ko), 영어(en), 중국어(zh) UI/에러/플레이스홀더 메시지.
 * Code.gs 및 Sidebar.html에서 사용하는 모든 사용자 대면 문구를 계층 구조로 정의.
 */

var I18N_LOCALE_KEY = 'LOCALE';

/**
 * ko / en / zh 메시지 팩. 계층: ui, alerts, placeholders, messages, errors
 */
var MESSAGES = {
  ko: {
    ui: {
      tabs: {
        manual: '직접 입력',
        paste: '통째로 붙여넣기',
        scan: '인용 검사',
        style: '서식 설정'
      },
      menu: {
        title: '인용 관리 메뉴',
        openSidebar: '사이드바 열기',
        clearData: '데이터 전체 초기화'
      },
      manual: {
        title: '논문/도서 제목',
        author: '저자 (세미콜론 \';\'으로 구분)',
        publisher: '출판사 / 학회 / 저널명',
        volume: '권 (Volume)',
        issue: '호 (Issue)',
        page: '페이지 (Page)',
        pagePrefix: 'p./pp. 표시',
        year: '연도',
        style: '양식'
      },
      paste: {
        shortLabel: '본문 인용구 제안 (수정 가능)',
        fullLabel: '전체 참고문헌 정보 (Full)',
        analyzeBtn: '지능형 분석 및 인용구 생성',
        yearLabel: '연도 (분석 결과)',
        btn_multi_insert: '선택된 문헌들 일괄 삽입',
        multiCitationWrapHint: '3개 이상 인용 시 본문이 길어져 줄바꿈이 될 수 있습니다.'
      },
      style: {
        description: '참고문헌 목록의 학회별 서식(Style Settings)을 설정합니다. DocumentProperties에 저장되어 문서 전체에 즉시 반영됩니다.',
        preset: '스타일 프리셋',
        italicsTitle: '논문/도서 제목 이탤릭체',
        italicsJournal: '저널명 이탤릭체',
        korItalicsJournal: '국문 저널명 이탤릭체 (해제 시 국문은 적용 안 함)',
        authorSeparator: '저자 구분자',
        etAlThreshold: 'et al./외 기준 (인)',
        lineSpacing: '줄 간격 (pt)',
        hangingIndent: '매달린 들여쓰기 (pt)',
        indentFirstLine: '첫 줄 들여쓰기 (pt)',
        headerFontSize: '헤더 글꼴 크기',
        headerAlignment: '헤더 정렬',
        yearBracket: '연도 괄호',
        saveStyle: '설정 저장',
        separatorComma: ', (쉼표)',
        separatorAmp: ' & (앰퍼샌드)',
        separatorSemi: '; (세미콜론)',
        alignLeft: '왼쪽',
        alignCenter: '가운데',
        alignRight: '오른쪽',
        bracketParen: '( )',
        bracketSquare: '[ ]',
        bracketNone: '없음',
        presetAPA: 'APA 7th',
        presetMLA: 'MLA',
        presetCustom: 'Custom',
        btn_finalize: '최종본 전환',
        btn_sort: '참고문헌 정렬 및 서식 적용'
      },
      scan: {
        description: '문서 내 임시 인용 (예: (Kim, 2024), [Kim24])을 감지하고 실제 문헌과 연결합니다.',
        scanBtn: '문서 스캔',
        applyMapping: '매핑 적용',
        recommended: '추천된 항목',
        unverified: '미확인 항목',
        skip: '대체하지 않음(Skip)',
        needLink: '⚠ 문헌 연결 필요',
        points: '점'
      },
      common: {
        insertHint: '삽입 위치: 본문에서 원하는 곳을 클릭한 뒤 버튼을 누르세요.',
        insertAndSave: '본문에 삽입 및 저장',
        saveOnly: '저장만 하기',
        generateBibliography: '참고문헌 목록 생성',
        clearAll: '데이터 전체 초기화',
        permissionHint: '처음 사용 시 \'권한 허용\' 창이 뜨면 허용해 주세요.',
        previewWaiting: '알림: 대기 중...'
      }
    },
    alerts: {
      requiredFields: '필수 정보를 입력하세요.',
      enterAllData: '데이터를 모두 입력하세요.',
      pasteFullFirst: '전체 참고문헌 정보(Full)를 먼저 붙여넣으세요.',
      confirmClear: '모든 데이터를 삭제하시겠습니까?',
      cleared: '초기화되었습니다.'
    },
    placeholders: {
      title: '제목을 입력하세요',
      author: 'Li, A; Smith, B',
      publisher: '출판 정보',
      volume: '25',
      issue: '6',
      page: '10-25 또는 110-120',
      year: '2024',
      shortCit: '예: (Smith, 2024) — 분석 버튼으로 자동 생성',
      fullPaste: '전체 내용을 붙여넣으세요.',
      pasteYear: '분석 후 자동 입력'
    },
    messages: {
      saving: '⏳ 저장 중...',
      processing: '⏳ 처리 중...',
      analyzing: '⏳ 분석 중...',
      generating: '⏳ 목록 생성 중...',
      scanning: '⏳ 문서 스캔 중...',
      applying: '⏳ 치환 적용 중...',
      styleSaved: '✅ 서식 설정이 저장되었습니다. 참고문헌 생성 시 적용됩니다.',
      analysisDone: '✅ 분석 완료. 본문 인용구 제안란을 확인·수정 후 저장하세요.',
      analysisFailed: '분석에 실패했습니다. 본문 인용구를 수동으로 입력해 주세요.',
      resultError: '⚠️ 결과 처리 중 오류가 발생했습니다. 본문 인용구를 수동으로 입력해 주세요.',
      insertAtCursor: '✅ 커서(선택) 위치에 삽입 및 저장 완료!',
      insertAtEnd: '✅ 본문 맨 끝에 삽입 및 저장 완료. 문서 맨 아래를 확인하세요.',
      multiInsertDone: '{count}개의 문헌이 다중 인용으로 삽입되었습니다.',
      saveDone: '✅ 저장 완료!',
      generateDone: '✅ 참고문헌 {count}개 생성 완료! (마지막 페이지)',
      generateNoItems: '⚠️ {message} 먼저 \'저장만 하기\' 또는 \'본문에 삽입 및 저장\'으로 항목을 추가하세요.',
      generateNoData: '⚠️ 저장된 참고문헌이 없습니다. 먼저 항목을 추가하세요.',
      clearDone: '✅ 초기화 완료',
      noDummies: '⚠️ 감지된 임시 인용이 없습니다.',
      noCitationsForScan: '⚠️ 저장된 문헌이 없습니다. 먼저 \'직접 입력\' 또는 \'통째로 붙여넣기\'로 문헌을 추가하세요.',
      scanDone: '✅ {count}개 인용 후보 감지됨',
      scanWithSections: ' (추천 {rec} / 미확인 {unv})',
      scanUnverifiedHint: '. 미확인 항목에 문헌을 연결해 주세요.',
      scanSelectHint: '. 문헌을 선택하고 \'매핑 적용\'을 누르세요.',
      mappingApplyFirst: '⚠️ 먼저 \'문서 스캔\'을 실행하세요.',
      mappingSelectItems: '⚠️ 매핑할 항목을 선택해 주세요.',
      mappingDone: '✅ {count}개 치환 완료! 참고문헌 목록 생성 시 반영됩니다.',
      finalizing: '색상 최적화 중...',
      finalizeDone: '{count}개 변경 완료',
      sorting: '문헌 정렬 및 내어쓰기 적용 중...',
      sortDone: '참고문헌이 저자순으로 정렬되고 서식이 적용되었습니다.',
      sortFormatDone: '정렬 및 서식 적용 완료',
      sortKoFirstDone: '국문 우선 정렬 및 서식 적용 완료',
      errorPrefix: '❌ 오류: ',
      errorServer: '❌ 오류(서버 연결): ',
      serverErrorPaste: ' 본문 인용구를 수동으로 입력해 주세요.'
    },
    errors: {
      noDataToSave: '저장할 데이터가 없습니다.',
      noDataToInsert: '삽입할 데이터가 없습니다.',
      noTextToParse: '입력된 텍스트가 없습니다.',
      parseError: '파싱 중 오류가 발생했습니다. 본문 인용구를 수동으로 입력해 주세요. ({detail})',
      bibliographyExists: '이미 \'참고문헌\' 섹션이 존재합니다. 기존 참고문헌 섹션을 지우거나 \'참고문헌\' 제목을 잠시 변경한 후 다시 시도해 주세요.',
      noCitationsStored: '저장된 참고문헌이 없습니다.',
      citationDataError: '저장 데이터 오류.',
      noMappingData: '매핑 데이터가 없습니다.',
      noTitle: '(제목 없음)',
      noContent: '(내용 없음)',
      operationFailed: '작업 중 오류가 발생했습니다. ({detail})'
    }
  },
  en: {
    ui: {
      tabs: { manual: 'Manual Entry', paste: 'Paste Full', scan: 'Citation Check', style: 'Style Settings' },
      menu: { title: 'Citation Menu', openSidebar: 'Open Sidebar', clearData: 'Clear All Data' },
      manual: {
        title: 'Title',
        author: 'Authors (separate with \';\')',
        publisher: 'Publisher / Journal',
        volume: 'Volume',
        issue: 'Issue',
        page: 'Page',
        pagePrefix: 'p./pp. prefix',
        year: 'Year',
        style: 'Style'
      },
      paste: {
        shortLabel: 'In-text citation suggestion (editable)',
        fullLabel: 'Full reference (Full)',
        analyzeBtn: 'Analyze & suggest citation',
        yearLabel: 'Year (from analysis)',
        btn_multi_insert: 'Insert Selected as Multi-cite',
        multiCitationWrapHint: 'With 3 or more citations, the in-text citation may become long and wrap to a new line.'
      },
      style: {
        description: 'Set bibliography style. Stored in DocumentProperties and applied when generating.',
        preset: 'Style preset',
        italicsTitle: 'Italicize title',
        italicsJournal: 'Italicize journal',
        korItalicsJournal: 'Italicize Korean journal (uncheck to skip)',
        authorSeparator: 'Author separator',
        etAlThreshold: 'et al. threshold (authors)',
        lineSpacing: 'Line spacing (pt)',
        hangingIndent: 'Hanging indent (pt)',
        indentFirstLine: 'First line indent (pt)',
        headerFontSize: 'Header font size',
        headerAlignment: 'Header alignment',
        yearBracket: 'Year bracket',
        saveStyle: 'Save settings',
        separatorComma: ', (comma)',
        separatorAmp: ' & (ampersand)',
        separatorSemi: '; (semicolon)',
        alignLeft: 'Left',
        alignCenter: 'Center',
        alignRight: 'Right',
        bracketParen: '( )',
        bracketSquare: '[ ]',
        bracketNone: 'None',
        presetAPA: 'APA 7th',
        presetMLA: 'MLA',
        presetCustom: 'Custom',
        btn_finalize: 'Finalize',
        btn_sort: 'Sort & Format Bibliography'
      },
      scan: {
        description: 'Detect in-text citations (e.g. (Kim, 2024)) and link to references.',
        scanBtn: 'Scan document',
        applyMapping: 'Apply mapping',
        recommended: 'Recommended',
        unverified: 'Unresolved',
        skip: 'Skip',
        needLink: '⚠ Link needed',
        points: 'pts'
      },
      common: {
        insertHint: 'Click in the document where you want to insert, then click the button.',
        insertAndSave: 'Insert & save',
        saveOnly: 'Save only',
        generateBibliography: 'Generate bibliography',
        clearAll: 'Clear all data',
        permissionHint: 'Please allow permission when prompted.',
        previewWaiting: 'Waiting...'
      }
    },
    alerts: {
      requiredFields: 'Please enter required fields.',
      enterAllData: 'Please enter all data.',
      pasteFullFirst: 'Please paste full reference text first.',
      confirmClear: 'Delete all data?',
      cleared: 'Cleared.'
    },
    placeholders: {
      title: 'Enter title',
      author: 'Li, A; Smith, B',
      publisher: 'Publisher',
      volume: '25',
      issue: '6',
      page: '10-25 or 110-120',
      year: '2024',
      shortCit: 'e.g. (Smith, 2024) — auto from analysis',
      fullPaste: 'Paste full reference here.',
      pasteYear: 'Filled after analysis'
    },
    messages: {
      saving: 'Saving...',
      processing: 'Processing...',
      analyzing: 'Analyzing...',
      generating: 'Generating...',
      scanning: 'Scanning...',
      applying: 'Applying...',
      styleSaved: 'Style settings saved.',
      analysisDone: 'Analysis done. Review and edit the citation suggestion, then save.',
      analysisFailed: 'Analysis failed. Enter citation manually.',
      resultError: 'Error processing result. Enter citation manually.',
      insertAtCursor: 'Inserted at cursor.',
      insertAtEnd: 'Inserted at end of document.',
      multiInsertDone: '{count} citations inserted as a combined group.',
      saveDone: 'Saved.',
      generateDone: 'Bibliography ({count} items) generated.',
      generateNoItems: '{message} Add items first.',
      generateNoData: 'No references stored. Add items first.',
      clearDone: 'Cleared.',
      noDummies: 'No in-text citations detected.',
      noCitationsForScan: 'No references stored. Add references first.',
      scanDone: '{count} citation(s) detected',
      scanWithSections: ' (recommended: {rec}, unresolved: {unv})',
      scanUnverifiedHint: '. Link unresolved items.',
      scanSelectHint: '. Select references and apply mapping.',
      mappingApplyFirst: 'Run document scan first.',
      mappingSelectItems: 'Select at least one item to map.',
      mappingDone: '{count} replacement(s) applied.',
      finalizing: 'Optimizing colors...',
      finalizeDone: '{count} change(s) completed',
      sorting: 'Sorting and applying hanging indents...',
      sortDone: 'Bibliography has been sorted and formatted.',
      sortFormatDone: 'Sort and format applied.',
      sortKoFirstDone: 'Korean-first sort and format applied.',
      errorPrefix: 'Error: ',
      errorServer: 'Server error: ',
      serverErrorPaste: ' Enter citation manually.'
    },
    errors: {
      noDataToSave: 'No data to save.',
      noDataToInsert: 'No data to insert.',
      noTextToParse: 'No text provided.',
      parseError: 'Parse error. Enter citation manually. ({detail})',
      bibliographyExists: 'A \'References\' section already exists. Remove it or rename the heading, then try again.',
      noCitationsStored: 'No references stored.',
      citationDataError: 'Invalid stored data.',
      noMappingData: 'No mapping data.',
      noTitle: '(No title)',
      noContent: '(No content)',
      operationFailed: 'An error occurred. ({detail})'
    }
  },
  zh: {
    ui: {
      tabs: { manual: '手动输入', paste: '整段粘贴', scan: '引用检查', style: '格式设置' },
      menu: { title: '引用管理菜单', openSidebar: '打开侧栏', clearData: '全部清除数据' },
      manual: {
        title: '论文/书名',
        author: '作者（用分号";"分隔）',
        publisher: '出版社/学会/期刊',
        volume: '卷 (Volume)',
        issue: '期 (Issue)',
        page: '页码 (Page)',
        pagePrefix: 'p./pp. 前缀',
        year: '年份',
        style: '格式'
      },
      paste: {
        shortLabel: '正文引用建议（可编辑）',
        fullLabel: '完整参考文献 (Full)',
        analyzeBtn: '智能分析并生成引用',
        yearLabel: '年份（分析结果）',
        btn_multi_insert: '一括插入选中的文献',
        multiCitationWrapHint: '3个及以上引用时，正文引用会变长，可能会换行。'
      },
      style: {
        description: '设置参考文献格式，保存于 DocumentProperties，生成时应用。',
        preset: '格式预设',
        italicsTitle: '论文/书名斜体',
        italicsJournal: '期刊名斜体',
        korItalicsJournal: '韩文期刊斜体（取消则不应用）',
        authorSeparator: '作者分隔符',
        etAlThreshold: 'et al./等 阈值（人）',
        lineSpacing: '行距 (pt)',
        hangingIndent: '悬挂缩进 (pt)',
        indentFirstLine: '首行缩进 (pt)',
        headerFontSize: '标题字号',
        headerAlignment: '标题对齐',
        yearBracket: '年份括号',
        saveStyle: '保存设置',
        separatorComma: '，（逗号）',
        separatorAmp: ' &（和号）',
        separatorSemi: '；（分号）',
        alignLeft: '左',
        alignCenter: '居中',
        alignRight: '右',
        bracketParen: '( )',
        bracketSquare: '[ ]',
        bracketNone: '无',
        presetAPA: 'APA 7th',
        presetMLA: 'MLA',
        presetCustom: 'Custom',
        btn_finalize: '最终版转换',
        btn_sort: '参考文献排序与格式化'
      },
      scan: {
        description: '检测正文中的临时引用（如 (Kim, 2024)）并链接到文献。',
        scanBtn: '扫描文档',
        applyMapping: '应用映射',
        recommended: '推荐',
        unverified: '未确认',
        skip: '不替换(Skip)',
        needLink: '⚠ 需链接文献',
        points: '分'
      },
      common: {
        insertHint: '在文档中点击插入位置后，再点击按钮。',
        insertAndSave: '插入正文并保存',
        saveOnly: '仅保存',
        generateBibliography: '生成参考文献',
        clearAll: '全部清除数据',
        permissionHint: '首次使用请允许权限。',
        previewWaiting: '等待中...'
      }
    },
    alerts: {
      requiredFields: '请填写必填项。',
      enterAllData: '请填写全部数据。',
      pasteFullFirst: '请先粘贴完整参考文献。',
      confirmClear: '确定删除全部数据？',
      cleared: '已清除。'
    },
    placeholders: {
      title: '输入标题',
      author: 'Li, A; Smith, B',
      publisher: '出版信息',
      volume: '25',
      issue: '6',
      page: '10-25 或 110-120',
      year: '2024',
      shortCit: '例：(Smith, 2024) — 由分析自动生成',
      fullPaste: '在此粘贴完整内容。',
      pasteYear: '分析后自动填入'
    },
    messages: {
      saving: '保存中...',
      processing: '处理中...',
      analyzing: '分析中...',
      generating: '生成中...',
      scanning: '扫描中...',
      applying: '应用替换中...',
      styleSaved: '格式设置已保存。',
      analysisDone: '分析完成。请确认或修改引用建议后保存。',
      analysisFailed: '分析失败。请手动输入引用。',
      resultError: '结果处理出错。请手动输入引用。',
      insertAtCursor: '已插入到光标位置。',
      insertAtEnd: '已插入到文档末尾。',
      multiInsertDone: '已将 {count} 个文献合并插入。',
      saveDone: '已保存。',
      generateDone: '已生成参考文献 {count} 条。',
      generateNoItems: '{message} 请先添加条目。',
      generateNoData: '暂无已保存的参考文献。请先添加条目。',
      clearDone: '已清除。',
      noDummies: '未检测到临时引用。',
      noCitationsForScan: '暂无已保存文献。请先添加。',
      scanDone: '检测到 {count} 个引用候选',
      scanWithSections: '（推荐 {rec} / 未确认 {unv}）',
      scanUnverifiedHint: '。请为未确认项链接文献。',
      scanSelectHint: '。选择文献后点击「应用映射」。',
      mappingApplyFirst: '请先执行「文档扫描」。',
      mappingSelectItems: '请至少选择一项进行映射。',
      mappingDone: '已替换 {count} 处。',
      finalizing: '颜色优化中...',
      finalizeDone: '已更改 {count} 处',
      sorting: '正在排序并应用悬挂缩进...',
      sortDone: '参考文献已按作者排序并完成格式化。',
      sortFormatDone: '排序与格式已应用。',
      sortKoFirstDone: '韩文优先排序与格式已应用。',
      errorPrefix: '错误：',
      errorServer: '服务器错误：',
      serverErrorPaste: ' 请手动输入引用。'
    },
    errors: {
      noDataToSave: '没有可保存的数据。',
      noDataToInsert: '没有可插入的数据。',
      noTextToParse: '未输入文本。',
      parseError: '解析出错。请手动输入引用。（{detail}）',
      bibliographyExists: '已存在「参考文献」节。请删除或暂时修改标题后再试。',
      noCitationsStored: '没有已保存的参考文献。',
      citationDataError: '存储数据异常。',
      noMappingData: '没有映射数据。',
      noTitle: '（无标题）',
      noContent: '（无内容）',
      operationFailed: '操作时发生错误。（{detail}）'
    }
  }
};

var SUPPORTED_LOCALES = ['ko', 'en', 'zh'];
var DEFAULT_LOCALE = 'ko';

/**
 * 사용자 로케일 감지.
 * 1) PropertiesService.getUserProperties()에 저장된 LOCALE 값 최우선
 * 2) 없으면 Session.getActiveUserLocale() 사용
 * 반환: 'ko' | 'en' | 'zh'
 */
function getUserLanguage() {
  try {
    var userProps = PropertiesService.getUserProperties();
    var saved = userProps.getProperty(I18N_LOCALE_KEY);
    if (saved && SUPPORTED_LOCALES.indexOf(saved) >= 0) return saved;
    var locale = Session.getActiveUserLocale();
    if (!locale || typeof locale !== 'string') return DEFAULT_LOCALE;
    locale = locale.toLowerCase();
    if (locale.indexOf('ko') === 0) return 'ko';
    if (locale.indexOf('zh') === 0) return 'zh';
    return 'en';
  } catch (e) {
    return DEFAULT_LOCALE;
  }
}

/**
 * 선택된 언어에 맞는 메시지 팩 반환.
 * @param {string} [lang] - 'ko' | 'en' | 'zh'. 생략 시 getUserLanguage() 결과 사용
 * @return {Object} 해당 언어의 MESSAGES[lang] (없으면 ko)
 */
function getLanguagePack(lang) {
  var key = lang && SUPPORTED_LOCALES.indexOf(lang) >= 0 ? lang : getUserLanguage();
  return MESSAGES[key] || MESSAGES[DEFAULT_LOCALE];
}

/**
 * 메시지 템플릿의 {key} 플레이스홀더를 치환.
 * @param {string} template - 예: "참고문헌 {count}개 생성 완료"
 * @param {Object} params - 예: { count: 5 }
 * @return {string}
 */
function formatMessage(template, params) {
  if (!template || !params) return template || '';
  var s = String(template);
  for (var key in params) {
    if (params.hasOwnProperty(key)) {
      s = s.replace(new RegExp('\\{' + key + '\\}', 'g'), String(params[key]));
    }
  }
  return s;
}
