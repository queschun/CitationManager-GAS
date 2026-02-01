🎓 Google Docs Citation Manager (et al. Automation)
"논문 작성 중 가장 번거로운 'et al.' 처리와 참고문헌 정렬을 자동화합니다."

이 프로젝트는 구글 문서(Google Docs)에서 논문을 작성할 때, 인용구 삽입부터 참고문헌 목록 생성까지의 과정을 혁신적으로 단축하기 위해 개발된 Google Apps Script(GAS) 기반 도구입니다.

✨ 핵심 기능 (Key Features)
지능형 'et al.' 자동화: 저자가 3인 이상일 경우, 국문 논문은 '외', 영문 논문은 **'et al.'**을 자동으로 판단하여 축약합니다.

이원화된 입력 모드:

직접 입력: 저자(세미콜론 구분), 제목, 연도, 출판사를 개별 필드로 입력.

통째로 붙여넣기: 구글 스칼라(Google Scholar) 등에서 복사한 인용 정보를 그대로 활용.

참고문헌 목록 원클릭 생성: 본문에 삽입된 데이터를 기반으로 문서 맨 마지막에 가나다/알파벳 순으로 정렬된 목록을 만듭니다.

스마트 섹션 관리: 참고문헌 생성 시 기존 목록을 삭제하고 새로 고침(Refresh)하여 페이지가 무한정 늘어나는 오류를 완벽히 해결했습니다.

데이터 영구 보관: PropertiesService를 활용해 문서 자체에 인용 데이터를 저장하므로, 문서를 껐다 켜도 데이터가 유지됩니다.

🛠 기술 스택 (Tech Stack)
Language: Google Apps Script (JavaScript)

Frontend: HTML5, CSS3 (Sidebar UI)

Integration: Google Docs API

Dev Tools: Cursor IDE, CLASP (Command Line Apps Script Projects)

🚀 시작하기 (Getting Started)
CLASP를 사용하는 경우
이 저장소를 클론합니다.

Bash
git clone https://github.com/사용자아이디/저장소이름.git
clasp login 후 clasp push를 통해 본인의 구글 스크립트 프로젝트로 전송합니다.

수동 설치 (Copy & Paste)
구글 문서에서 확장 프로그램 > Apps Script를 엽니다.

Code.gs와 Sidebar.html 파일의 내용을 복사하여 붙여넣습니다.

저장 후 구글 문서를 새로고침하면 상단에 **[인용 관리 메뉴]**가 나타납니다.

📝 사용 규칙 (Usage Rules)
저자 입력: 외국 저자의 성과 이름을 명확히 구분하기 위해 저자 간 구분자는 **세미콜론(;)**을 사용합니다.

예: Li, A; Smith, B

참고문헌 양식: 현재 APA 및 MLA 스타일의 기초 양식을 지원하며, 필요에 따라 코드를 수정하여 커스터마이징할 수 있습니다.

👤 Author
Chun, Seichul 

📄 License
이 프로젝트는 MIT License를 따릅니다. 누구나 자유롭게 수정하고 배포할 수 있습니다.