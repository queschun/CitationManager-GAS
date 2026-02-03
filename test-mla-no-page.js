// MLA 스타일 + 페이지 번호 없는 문헌 시뮬레이션 (Code.gs 로직 재현)
var item = {
  mode: 'manual',
  author: 'Smith, John',
  authorRaw: 'Smith, John',
  secondVal: '2020',
  year: '2020',
  title: 'Testing Without Pages',
  publisher: 'Test Press',
  page: '',
  style: 'MLA'
};

// 본문 인용 (getFormattedCitation MLA): page 없으면 (Author)만
var authorInText = 'Smith';
var page = (item.page || '').toString().trim();
var inText = page ? '(' + authorInText + ' ' + page + ')' : '(' + authorInText + ')';

// 참고문헌 (formatCitationEntry MLA): page 없으면 pp. 접두사·쉼표 없음
var authorStr = item.author;
var title = (item.title || '').toString().trim();
var year = (item.secondVal != null ? item.secondVal : item.year || '').toString().trim();
var pub = (item.publisher || '').toString().trim();
var mla = authorStr + '. "' + title + '." ';
if (pub) mla += pub + ', ';
mla += year + '.';

console.log('=== MLA no-page simulation ===');
console.log('In-text citation:   ', inText);
console.log('Bibliography entry:', mla);
console.log('');
console.log('Checks:');
console.log('  inText is (Author) only:', inText === '(Smith)');
console.log('  bibliography has no pp.:', mla.indexOf('pp.') === -1);
