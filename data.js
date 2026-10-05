/* 이 파일은 scripts/build_data.py가 자동으로 만들어요. 직접 고치지 말고 엑셀을 고친 뒤 다시 실행하세요.
   원본: data/회계법인_용어집.xlsx · 만든 시각: 2026-10-05 16:34 · 용어 45개 */
const CATEGORIES = [
  {
    "id": "slang",
    "name": "은어·줄임말",
    "color": "rose",
    "desc": "묻기 애매한 줄임말과 은어를 모았어요. 여기서 먼저 찾아보세요.",
    "emoji": "🔥",
    "sample": [
      "금조",
      "FYI",
      "F/U",
      "빈콩",
      "반콩"
    ],
    "count": 16
  },
  {
    "id": "process",
    "name": "업무 프로세스",
    "color": "orange",
    "desc": "어사인, 조서, 감사 절차처럼 일하는 순서와 관련된 말이에요.",
    "sample": [
      "어사인",
      "필드",
      "테일러링",
      "PbC",
      "필드워크"
    ],
    "count": 10
  },
  {
    "id": "org",
    "name": "조직·직급",
    "color": "tangerine",
    "desc": "직급과 역할 이름이에요. 누구에게 무엇을 물어볼지 알 수 있어요.",
    "sample": [
      "Associate",
      "Manager",
      "PM (Project Manager)",
      "인차지",
      "스태프"
    ],
    "count": 11
  },
  {
    "id": "client",
    "name": "감사 개념",
    "color": "yellow",
    "desc": "중요성, 위험평가처럼 감사 기준에서 나오는 개념이에요.",
    "sample": [
      "FSLI",
      "LSPM",
      "RoMM",
      "OM",
      "PM (Performance Materiality)"
    ],
    "count": 7
  },
  {
    "id": "tool",
    "name": "시스템·툴",
    "color": "gray",
    "desc": "매일 쓰는 사내 시스템과 프로그램 이름이에요.",
    "sample": [
      "Aura"
    ],
    "count": 1
  }
];

const TERMS = [
  {
    "id": "geumjo",
    "name": "금조",
    "category": "slang",
    "level": "week1",
    "definition": "금융기관 조회서의 줄임말. 감사 절차상 완전성 검토를 통해 식별된 금융기관에 발송하여 예금·차입금 등 금융정보를 외부에서 직접 확인하는 조회서. 1년 차 직원이 주로 담당하는 핵심 업무 중 하나.",
    "aliases": [
      "금융기관 조회서"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "byeonjo",
    "name": "변조",
    "category": "slang",
    "level": "month1",
    "definition": "변호사 조회서의 줄임말. 회사의 법률대리인 또는 담당 변호사에게 발송하여 소송·분쟁·우발사항 등을 확인하는 조회서.",
    "aliases": [
      "변호사 조회서"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "chaechaejo",
    "name": "채채조",
    "category": "slang",
    "level": "month1",
    "definition": "채권채무 조회서의 줄임말. 회사와 채권·채무 관계가 있는 거래처 등에 발송하여 관련 잔액을 확인하는 조회서.",
    "aliases": [
      "채권채무 조회서"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "assign",
    "name": "어사인",
    "category": "process",
    "level": "week1",
    "definition": "개인에게 배정된 감사 프로젝트 또는 업무 일정. 날짜별로 어느 클라이언트·프로젝트에 배정되어 있는지를 의미하며, 일정이 없으면 ‘어사인이 비었다’고 표현한다.",
    "aliases": [
      "Assign"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "field",
    "name": "필드",
    "category": "process",
    "level": "week1",
    "definition": "감사 대상인 클라이언트 회사 또는 해당 회사에서 업무를 수행하는 현장을 의미한다. 물리적인 장소와 프로젝트 양쪽 의미로 사용된다.",
    "aliases": [
      "Field"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "fieldwork",
    "name": "필드워크",
    "category": "process",
    "level": "month1",
    "definition": "클라이언트 회사 또는 프로젝트 현장에서 실질적인 감사 업무를 수행하는 기간 및 업무.",
    "aliases": [
      "Fieldwork",
      "Field work"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "tailoring",
    "name": "테일러링",
    "category": "process",
    "level": "week1",
    "definition": "전기 조서 등을 당기 감사에 맞게 수정·정리·세팅하는 작업. 연도, 금액, 기간, 설명 등을 당해 연도 기준으로 바꾸는 업무를 포함한다.",
    "aliases": [
      "Tailoring"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "ar",
    "name": "AR",
    "category": "process",
    "level": "month1",
    "definition": "Analytical Review/Analytical Procedure. 계정의 증감, 추세 및 눈에 띄는 변화를 확인하고 그 발생 원인과 근거를 분석하는 절차.",
    "aliases": [
      "Analytical Review",
      "Analytical Procedure",
      "분석적 절차"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "asis",
    "name": "AS-IS",
    "category": "process",
    "level": "month1",
    "definition": "특정 항목이 현재 실제로 처리·계산·표시되어 있는 상태 또는 금액. 오류 수정이나 개선 논의 시 TO-BE와 대비하여 사용한다.",
    "aliases": [
      "애즈이즈"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "tobe",
    "name": "TO-BE",
    "category": "process",
    "level": "month1",
    "definition": "검토·수정 후 올바르게 되어야 하는 상태 또는 금액. 현재 상태인 AS-IS와 비교하여 차이를 분석할 때 사용한다.",
    "aliases": [
      "투비"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "fyi",
    "name": "FYI",
    "category": "slang",
    "level": "week1",
    "definition": "For Your Information의 약자. 직접적인 업무 수행을 요청하기보다는 참고 또는 정보 공유 목적으로 자료나 메일을 전달할 때 사용한다.",
    "aliases": [
      "For Your Information"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "fu",
    "name": "F/U",
    "category": "slang",
    "level": "week1",
    "definition": "Follow-up의 약자. 최초 업무 수행 이후 추가적으로 확인·보완·조치가 필요한 후속 업무. 필드 철수 이후 남은 일을 처리하는 경우에도 자주 쓴다.",
    "aliases": [
      "FU",
      "Follow-up",
      "팔로업"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "pbc",
    "name": "PbC",
    "category": "process",
    "level": "week1",
    "definition": "Provided by Client. 감사 업무를 위해 클라이언트가 제공해야 하거나 제공한 자료를 의미하며, 요청 자료 목록을 PbC List라고 부르는 경우가 많다.",
    "aliases": [
      "PBC",
      "Provided by Client",
      "PbC List"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "refer",
    "name": "레퍼",
    "category": "process",
    "level": "month1",
    "definition": "서로 연관된 조서·재무제표·주석·금액 등이 상호 참조되거나 일치하는지 확인하는 것. 연결되어야 할 숫자가 불일치하면 ‘레퍼가 안 된다’고 표현한다.",
    "aliases": [
      "Refer",
      "Cross-reference"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "wtt",
    "name": "WTT",
    "category": "process",
    "level": "month1",
    "definition": "Walkthrough Test. 하나의 거래나 사례를 선정해 거래 발생부터 회계 처리까지 회사의 프로세스를 따라가며 업무 흐름과 내부통제의 설계·구현 등을 확인하는 절차.",
    "aliases": [
      "Walkthrough Test",
      "Walkthrough",
      "워크스루"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "aura",
    "name": "Aura",
    "category": "tool",
    "level": "week1",
    "definition": "삼일 감사에서 감사조서를 작성·관리할 때 사용하는 감사 조서 시스템.",
    "aliases": [
      "오라"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "binkong",
    "name": "빈콩",
    "category": "slang",
    "level": "week1",
    "definition": "Aura에서 전기 조서가 포워딩된 후 아직 당기 작성·수정이 이루어지지 않은 상태를 의미하는 은어. 조서 제목 옆 원이 ○ 상태인 데서 유래한다.",
    "aliases": [],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "bankong",
    "name": "반콩",
    "category": "slang",
    "level": "week1",
    "definition": "Aura에서 조서에 일부 내용이 작성·수정되었지만 Prepared 처리가 완료되지 않은 상태를 의미하는 은어. ◐처럼 반쯤 채워진 표시에서 유래한다.",
    "aliases": [],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "kongjjikgi",
    "name": "콩찍기",
    "category": "slang",
    "level": "week1",
    "definition": "Aura에서 조서 작성을 완료한 뒤 Prepared 버튼을 눌러 작성 완료 상태로 만드는 것. 원 표시가 ●로 채워지는 모습에서 유래한 은어.",
    "aliases": [
      "콩을 찍다",
      "콩 찍다",
      "콩 찍"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "galmaegi",
    "name": "갈매기",
    "category": "slang",
    "level": "month1",
    "definition": "Prepared된 조서에 대해 담당 리뷰어가 검토를 완료한 상태를 나타내는 체크 표시를 지칭하는 은어.",
    "aliases": [],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "ap",
    "name": "AP",
    "category": "slang",
    "level": "week1",
    "definition": "아모레퍼시픽 빌딩을 지칭하는 내부 표현. 신용산에 위치한 삼일 사무실을 말하며, 필드가 아닌 사무실 근무 여부를 표현할 때 사용한다.",
    "aliases": [
      "아모레퍼시픽 빌딩"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "associate",
    "name": "Associate",
    "category": "org",
    "level": "week1",
    "definition": "삼일 직급체계의 실무진 시작 단계에 해당하는 직급.",
    "aliases": [
      "어쏘"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "seniorassociate",
    "name": "Senior Associate",
    "category": "org",
    "level": "month1",
    "definition": "Associate의 상위 직급. 프로젝트에서 보다 독립적으로 업무를 수행하고 하위 연차 업무를 검토하는 역할을 맡을 수 있다.",
    "aliases": [
      "SA",
      "시니어"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "manager",
    "name": "Manager",
    "category": "org",
    "level": "week1",
    "definition": "Senior Associate의 상위 관리 직급. 감사 프로젝트의 실무 관리·검토 등을 담당한다.",
    "aliases": [
      "매니저"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "seniormanager",
    "name": "Senior Manager",
    "category": "org",
    "level": "month1",
    "definition": "Manager의 상위 직급. 규모가 큰 프로젝트의 관리나 PM 역할 등을 수행할 수 있다.",
    "aliases": [
      "SM",
      "시니어 매니저"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "director",
    "name": "Director",
    "category": "org",
    "level": "month1",
    "definition": "Senior Manager의 상위 리더십 직급. 감사 프로젝트에서 PM 등의 역할을 수행할 수 있다.",
    "aliases": [
      "디렉터"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "partner",
    "name": "Partner",
    "category": "org",
    "level": "month1",
    "definition": "법인의 파트너 직급. 감사 프로젝트에서 EL 등 최종 책임 역할을 맡을 수 있으며, EP·NEP는 파트너 구분에 사용되는 내부 표기.",
    "aliases": [
      "파트너",
      "EP",
      "NEP"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "el",
    "name": "EL",
    "category": "org",
    "level": "month1",
    "definition": "Engagement Leader. 감사 프로젝트의 업무수행이사 역할로, 통상 Partner가 담당하며 프로젝트의 최상위 책임자 역할을 수행한다.",
    "aliases": [
      "Engagement Leader"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "pmproject",
    "name": "PM (Project Manager)",
    "category": "org",
    "level": "week1",
    "definition": "감사 프로젝트의 전반적인 일정과 업무를 관리하는 역할. 보통 Director나 Senior Manager가 맡고, 규모가 작은 프로젝트에서는 Manager가 맡기도 한다.",
    "aliases": [
      "PM",
      "Project Manager"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "incharge",
    "name": "인차지",
    "category": "org",
    "level": "week1",
    "definition": "감사 현장에서 실무진을 관리하고 주요 업무 진행을 책임지는 현장 실무 책임자.",
    "aliases": [
      "In-charge",
      "In charge"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "km",
    "name": "KM",
    "category": "org",
    "level": "month1",
    "definition": "Key Member. 프로젝트의 핵심 실무 구성원을 의미하며, 실무에서는 인차지를 가리키거나 유사한 역할을 뜻하는 경우가 있다.",
    "aliases": [
      "Key Member"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "staff",
    "name": "스태프",
    "category": "org",
    "level": "week1",
    "definition": "감사 프로젝트에서 실제 감사 절차와 조서 작성 등을 수행하는 실무 구성원.",
    "aliases": [
      "Staff"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "gaduri",
    "name": "가두리",
    "category": "slang",
    "level": "week1",
    "definition": "감사 업무 시 팀원들을 한 회의실 또는 특정 공간에 모아두고 함께 업무를 수행하게 하는 방식을 의미하는 사내 은어.",
    "aliases": [],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "geunduri",
    "name": "근두리",
    "category": "slang",
    "level": "month1",
    "definition": "팀원들을 한 공간에 모두 모으지는 않지만, 필요할 때 바로 소통하거나 확인할 수 있도록 서로 가까운 위치에 두고 업무를 수행하게 하는 방식을 의미하는 사내 은어.",
    "aliases": [],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "coreduri",
    "name": "코어두리",
    "category": "slang",
    "level": "month1",
    "definition": "팀원들을 특정 코어(Core) 또는 정해진 구역을 중심으로 모아두고 업무를 수행하게 하는 방식을 의미하는 사내 은어.",
    "aliases": [],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "gwangduri",
    "name": "광두리",
    "category": "slang",
    "level": "month1",
    "definition": "팀원들을 한 장소나 가까운 범위에 모아두지 않고 비교적 넓은 범위에 분산시켜 각자 업무를 수행하게 하는 방식을 의미하는 사내 은어.",
    "aliases": [],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "naraebi",
    "name": "나래비",
    "category": "slang",
    "level": "month1",
    "definition": "여러 항목이나 자료를 순서대로 쭉 나열하거나 늘어놓는다는 의미로 쓰는 구어적 표현.",
    "aliases": [],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "사용자 제공 사내 용례",
    "views": 0
  },
  {
    "id": "fsli",
    "name": "FSLI",
    "category": "client",
    "level": "month1",
    "definition": "Financial Statement Line Item. 재무제표를 구성하는 개별 계정·항목 단위. 실무에서는 각 직원에게 특정 FSLI가 배정되고, 해당 항목의 위험과 감사절차를 수행하는 식으로 사용된다.",
    "aliases": [
      "Financial Statement Line Item"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "https://note.jp.pwc.com/n/n79885b5314a1?hl=en",
    "views": 0
  },
  {
    "id": "lspm",
    "name": "LSPM",
    "category": "client",
    "level": "later",
    "definition": "Likely Source of Potential Misstatement. 프로세스상 잠재적 왜곡표시가 발생할 가능성이 있는 구체적인 원천·지점. 쉽게 말해 ‘이 프로세스에서 어디서, 어떻게 잘못될 수 있는가’를 식별하는 개념.",
    "aliases": [
      "Likely Source of Potential Misstatement"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "https://www.pwc.com/gr/en/events/assets/internal-controls.pdf",
    "views": 0
  },
  {
    "id": "romm",
    "name": "RoMM",
    "category": "client",
    "level": "later",
    "definition": "Risk of Material Misstatement. 재무제표가 감사 수행 전에 중요하게 왜곡표시되어 있을 위험. 재무제표 전체 수준 또는 개별 assertion 수준에서 식별·평가하고, 그 결과에 따라 후속 감사절차를 설계한다.",
    "aliases": [
      "ROMM",
      "Risk of Material Misstatement"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "https://www.iaasb.org/publications/isa-315-revised-2019-identifying-and-assessing-risks-material-misstatement",
    "views": 0
  },
  {
    "id": "om",
    "name": "OM",
    "category": "client",
    "level": "later",
    "definition": "Overall Materiality. 재무제표 전체 수준의 중요성 금액. 재무제표 이용자의 의사결정에 영향을 줄 수 있는 왜곡표시를 판단하기 위한 기본 중요성 기준.",
    "aliases": [
      "Overall Materiality",
      "중요성"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "https://www.iaasb.org/consultations-projects/materiality-planning-and-performing-audit-and-evaluation-misstatements-identified-during-audit-isa",
    "views": 0
  },
  {
    "id": "pmperformance",
    "name": "PM (Performance Materiality)",
    "category": "client",
    "level": "later",
    "definition": "Performance Materiality. 감사절차를 설계·수행할 때 사용하는 수행중요성 금액으로, 일반적으로 OM보다 낮게 설정하여 개별 오류의 누적 위험을 낮추는 데 사용한다.",
    "aliases": [
      "PM",
      "Performance Materiality",
      "수행중요성"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "https://www.iaasb.org/consultations-projects/materiality-planning-and-performing-audit-and-evaluation-misstatements-identified-during-audit-isa",
    "views": 0
  },
  {
    "id": "dm",
    "name": "DM",
    "category": "client",
    "level": "later",
    "definition": "De Minimis. 감사 중 발견한 차이·왜곡표시 가운데 너무 사소하여 SUM에 누적하지 않아도 되는 기준선으로 실무에서 사용한다. 다만 금액이 작더라도 성격상 중요한 사항은 별도 판단이 필요하다.",
    "aliases": [
      "De Minimis"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "https://www.iaasb.org/consultations-projects/materiality-planning-and-performing-audit-and-evaluation-misstatements-identified-during-audit-isa",
    "views": 0
  },
  {
    "id": "sum",
    "name": "SUM",
    "category": "client",
    "level": "later",
    "definition": "Summary of Uncorrected Misstatements. 감사 과정에서 발견했지만 회사가 수정하지 않은 왜곡표시를 모아 관리·평가하는 내역. 감사 종료 단계에서 누적 영향과 중요성을 검토한다.",
    "aliases": [
      "Summary of Uncorrected Misstatements"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "https://www.pwc.com/jg/en/publications/audit-comm-effectiveness-what-works-best-2011.pdf",
    "views": 0
  },
  {
    "id": "minorpass",
    "name": "Minor pass",
    "category": "slang",
    "level": "later",
    "definition": "감사 실무에서 ‘회계처리상 오류는 맞지만 금액이나 영향이 중요하지 않아 수정 요구 없이 넘어가는 것’을 뜻하는 표현. 정식 감사기준 용어라기보다는 현업에서 쓰는 실무 표현이며, 금액이 작더라도 질적으로 중요한 사항은 단순히 패스하면 안 된다.",
    "aliases": [
      "마이너 패스"
    ],
    "createdBy": "익명 선배",
    "createdAt": "2026-10-03",
    "updatedAt": "2026-10-03",
    "reviewStatus": "pending",
    "reviewedBy": null,
    "source": "https://cpa-map.com/post-2124/",
    "views": 0
  }
];

const DETAILS = {
  "geumjo": {
    "examples": [
      {
        "where": "",
        "text": "선생님, 금조 완전성 검토 마무리되었나요?",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "금조 회수 현황 업데이트해주세요.",
        "by": "익명 선배"
      }
    ],
    "related": [],
    "keywords": [
      "조회서",
      "완전성 검토",
      "외부조회",
      "발송",
      "회수"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "byeonjo": {
    "examples": [
      {
        "where": "",
        "text": "변조 발송 대상 변호사 리스트 확인해주세요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "변조 회신 아직 안 온 곳 있어요?",
        "by": "익명 선배"
      }
    ],
    "related": [],
    "keywords": [
      "소송",
      "우발부채",
      "외부조회",
      "법무법인"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "chaechaejo": {
    "examples": [
      {
        "where": "",
        "text": "채채조 회신 안 온 거래처 다시 팔로업해주세요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "채채조 대상 리스트 정리됐나요?",
        "by": "익명 선배"
      }
    ],
    "related": [],
    "keywords": [
      "매출채권",
      "매입채무",
      "거래처",
      "외부조회"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "assign": {
    "examples": [
      {
        "where": "",
        "text": "이번 주 어사인 있어?",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "다음 주 어사인 컨플릭 났어.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "나 다음 주 어사인 비었어.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "field"
    ],
    "keywords": [
      "배정",
      "스케줄",
      "Pooling",
      "프로젝트",
      "어사인 컨플릭"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "field": {
    "examples": [
      {
        "where": "",
        "text": "이번 주 필드 어디였어? — 나 이번 주 내내 삼성전자였어.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "다음 주 필드에서 뵙겠습니다.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "fieldwork",
      "assign",
      "ap"
    ],
    "keywords": [
      "클라이언트",
      "현장감사"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "fieldwork": {
    "examples": [
      {
        "where": "",
        "text": "다음 주부터 필드워크 시작해요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "필드워크 끝나고 남은 건 F/U 하면 돼요.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "field",
      "fu"
    ],
    "keywords": [
      "현장감사",
      "클라이언트",
      "감사 프로젝트"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "tailoring": {
    "examples": [
      {
        "where": "",
        "text": "쌤, 현예금 조서들 테일러링 부탁해요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "이 조서 아직 작년 숫자라 테일러링부터 해야 돼요.",
        "by": "익명 선배"
      }
    ],
    "related": [],
    "keywords": [
      "조서",
      "포워딩",
      "전기 조서",
      "당기 조서",
      "세팅"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "ar": {
    "examples": [
      {
        "where": "",
        "text": "이번 주까지 AR 마무리해주세요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "매출 AR 하면서 전년 대비 증감 원인 확인해주세요.",
        "by": "익명 선배"
      }
    ],
    "related": [],
    "keywords": [
      "증감분석",
      "추세분석",
      "계정분석",
      "감사절차"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "asis": {
    "examples": [
      {
        "where": "",
        "text": "현재 임차보증금 현할차 AS-IS 금액 먼저 뽑아주세요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "일단 AS-IS가 어떻게 되어 있는지부터 봐주세요.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "tobe"
    ],
    "keywords": [
      "현재 상태",
      "오류",
      "수정",
      "차이분석"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "tobe": {
    "examples": [
      {
        "where": "",
        "text": "AS-IS랑 TO-BE 금액 비교해서 차이 정리해주세요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "TO-BE 기준으로 다시 계산해볼게요.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "asis"
    ],
    "keywords": [
      "목표 상태",
      "수정금액",
      "차이분석"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "fyi": {
    "examples": [
      {
        "where": "",
        "text": "FYI로 전달드립니다.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "아래 메일 FYI입니다.",
        "by": "익명 선배"
      }
    ],
    "related": [],
    "keywords": [
      "참고",
      "정보공유",
      "메일",
      "Forward"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "fu": {
    "examples": [
      {
        "where": "",
        "text": "이건 3분기 끝나고 팔로업해주세요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "미수령 PbC F/U 부탁드려요.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "pbc"
    ],
    "keywords": [
      "후속조치",
      "미결사항",
      "Open item"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "pbc": {
    "examples": [
      {
        "where": "",
        "text": "다음 월요일에 클라한테 던질 거니까 이번 주까지 PbC 리스트 수합해주세요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "이 PbC 아직 안 왔어요.",
        "by": "익명 선배"
      }
    ],
    "related": [],
    "keywords": [
      "요청자료",
      "클라이언트",
      "자료수합"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "refer": {
    "examples": [
      {
        "where": "",
        "text": "지금 주석 13번 금액이랑 16번 금액이 레퍼가 안 되네요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "이 숫자 어디로 레퍼돼요?",
        "by": "익명 선배"
      }
    ],
    "related": [],
    "keywords": [
      "참조",
      "재무제표",
      "주석",
      "금액 일치",
      "Tie-out"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "wtt": {
    "examples": [
      {
        "where": "",
        "text": "이번 주에 매출 프로세스 WTT 잡혀 있어요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "WTT용 샘플 하나 받아주세요.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "lspm"
    ],
    "keywords": [
      "내부통제",
      "프로세스",
      "통제",
      "인터뷰"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "aura": {
    "examples": [
      {
        "where": "",
        "text": "Aura에 조서 올려주세요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "전기 조서 Aura에서 포워딩해왔어요.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "binkong",
      "bankong",
      "kongjjikgi",
      "galmaegi"
    ],
    "keywords": [
      "감사조서",
      "Prepared",
      "Review"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "binkong": {
    "examples": [
      {
        "where": "",
        "text": "아직 빈콩인 조서들 먼저 확인해주세요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "이거 아직 빈콩이네.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "aura",
      "bankong",
      "kongjjikgi"
    ],
    "keywords": [
      "조서",
      "포워딩"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "bankong": {
    "examples": [
      {
        "where": "",
        "text": "반콩 상태인 조서들 이번 주 안에 마무리해주세요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "일단 반콩까지는 만들어놨어요.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "aura",
      "binkong",
      "kongjjikgi"
    ],
    "keywords": [
      "Prepared"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "kongjjikgi": {
    "examples": [
      {
        "where": "",
        "text": "선생님, 이번 주까지 조서 다 콩 찍어주세요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "이 조서 이제 콩 찍어도 돼요?",
        "by": "익명 선배"
      }
    ],
    "related": [
      "aura",
      "binkong",
      "bankong",
      "galmaegi"
    ],
    "keywords": [
      "Prepared",
      "조서 완료"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "galmaegi": {
    "examples": [
      {
        "where": "",
        "text": "이 조서는 갈매기까지 떴어요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "갈매기 안 뜬 조서만 다시 봐주세요.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "aura",
      "kongjjikgi"
    ],
    "keywords": [
      "Review",
      "Reviewer",
      "Prepared"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "ap": {
    "examples": [
      {
        "where": "",
        "text": "이번 주 AP야, 아니면 필드야?",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "오늘은 AP에서 일해.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "field"
    ],
    "keywords": [
      "사무실",
      "신용산"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "associate": {
    "examples": [
      {
        "where": "",
        "text": "이번 프로젝트 어쏘 몇 명이에요?",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "저 어쏘 1년 차예요.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "seniorassociate",
      "staff"
    ],
    "keywords": [
      "직급",
      "Junior"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "seniorassociate": {
    "examples": [
      {
        "where": "",
        "text": "이번 프로젝트 Senior Associate가 인차지예요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "SA 리뷰 먼저 받고 올려주세요.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "associate",
      "manager",
      "incharge"
    ],
    "keywords": [],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "manager": {
    "examples": [
      {
        "where": "",
        "text": "Manager 리뷰 후 수정사항 반영해주세요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "이 이슈는 매니저님께 먼저 말씀드릴게요.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "seniorassociate",
      "seniormanager",
      "pmproject"
    ],
    "keywords": [
      "Review"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "seniormanager": {
    "examples": [
      {
        "where": "",
        "text": "이번 프로젝트 PM은 SM이에요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "SM 리뷰까지 끝났대요.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "manager",
      "director",
      "pmproject"
    ],
    "keywords": [
      "직급"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "director": {
    "examples": [
      {
        "where": "",
        "text": "이번 프로젝트 PM은 Director님이에요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "Director 리뷰 일정 잡혔어요?",
        "by": "익명 선배"
      }
    ],
    "related": [
      "seniormanager",
      "partner",
      "pmproject"
    ],
    "keywords": [
      "직급"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "partner": {
    "examples": [
      {
        "where": "",
        "text": "이번 감사 EL은 ○○ 파트너님이에요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "파트너 리뷰 전까지 조서 정리해주세요.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "el",
      "director"
    ],
    "keywords": [],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "el": {
    "examples": [
      {
        "where": "",
        "text": "EL 리뷰 일정 확인해주세요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "이 이슈 EL까지 올라갔어요?",
        "by": "익명 선배"
      }
    ],
    "related": [
      "partner",
      "pmproject"
    ],
    "keywords": [
      "감사 프로젝트"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "pmproject": {
    "examples": [
      {
        "where": "",
        "text": "이번 프로젝트 PM 누구예요?",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "PM한테 먼저 확인하고 진행할게요.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "el",
      "director",
      "seniormanager",
      "manager",
      "incharge"
    ],
    "keywords": [
      "프로젝트 관리"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "incharge": {
    "examples": [
      {
        "where": "",
        "text": "이 이슈는 인차지한테 먼저 공유해주세요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "인차지 리뷰 받고 콩 찍어주세요.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "km",
      "staff",
      "pmproject"
    ],
    "keywords": [
      "현장책임자",
      "Senior"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "km": {
    "examples": [
      {
        "where": "",
        "text": "이번 프로젝트 KM이 누구예요?",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "KM한테 이슈 공유했어요?",
        "by": "익명 선배"
      }
    ],
    "related": [
      "incharge",
      "pmproject",
      "staff"
    ],
    "keywords": [],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "staff": {
    "examples": [
      {
        "where": "",
        "text": "이 계정은 스태프한테 배정해주세요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "스태프들 PbC 먼저 정리해주세요.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "associate",
      "km",
      "incharge"
    ],
    "keywords": [],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "gaduri": {
    "examples": [
      {
        "where": "",
        "text": "나 지금 가두리 중이라 못 나가.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "우리 이번 주 내내 가두리래.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "오늘도 회의실 가두리야?",
        "by": "익명 선배"
      }
    ],
    "related": [
      "fieldwork",
      "incharge",
      "geunduri",
      "coreduri",
      "gwangduri"
    ],
    "keywords": [
      "회의실",
      "팀원 배치"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "geunduri": {
    "examples": [
      {
        "where": "",
        "text": "가두리는 아닌데 근두리라 멀리는 못 가.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "오늘 근두리라 이 근처에서 일해야 돼.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "gaduri",
      "coreduri",
      "gwangduri",
      "fieldwork"
    ],
    "keywords": [
      "근거리",
      "팀원 배치"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "coreduri": {
    "examples": [
      {
        "where": "",
        "text": "가두리까진 아니고 코어두리래.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "우리 오늘 이 코어 안에서만 있으면 된대.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "gaduri",
      "geunduri",
      "gwangduri"
    ],
    "keywords": [
      "Core",
      "업무 공간",
      "팀원 배치"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "gwangduri": {
    "examples": [
      {
        "where": "",
        "text": "이번 주 광두리라 각자 편한 데서 일하면 돼.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "우리 팀 광두리라 오늘 얼굴 못 볼 수도 있어.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "gaduri",
      "geunduri",
      "coreduri",
      "fieldwork"
    ],
    "keywords": [
      "분산근무",
      "팀원 배치"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "naraebi": {
    "examples": [
      {
        "where": "",
        "text": "이 숫자들 나래비로 쭉 정리해주세요.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "계정별로 나래비 해놓으면 보기 편해.",
        "by": "익명 선배"
      }
    ],
    "related": [],
    "keywords": [
      "나열",
      "리스트업",
      "정리",
      "배열"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "fsli": {
    "examples": [
      {
        "where": "",
        "text": "이번에 매출 FSLI 누가 맡았어?",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "재고 FSLI 조서부터 먼저 봐주세요.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "romm",
      "lspm"
    ],
    "keywords": [
      "계정",
      "Scoping",
      "Assertion"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "lspm": {
    "examples": [
      {
        "where": "",
        "text": "매출 프로세스 LSPM 뭐 잡았어요?",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "이 컨트롤이 어떤 LSPM을 커버해요?",
        "by": "익명 선배"
      }
    ],
    "related": [
      "wtt"
    ],
    "keywords": [
      "내부통제",
      "Process",
      "Control",
      "What could go wrong"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "romm": {
    "examples": [
      {
        "where": "",
        "text": "이 FSLI에 잡힌 RoMM이 뭐예요?",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "매출 인식 쪽 RoMM이 높아서 절차를 더 가져가야 한대.",
        "by": "익명 선배"
      }
    ],
    "related": [],
    "keywords": [
      "Inherent Risk",
      "Control Risk",
      "Significant Risk",
      "Assertion"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "om": {
    "examples": [
      {
        "where": "",
        "text": "이번 회사 OM 얼마예요?",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "올해 실적 바뀌어서 OM 다시 계산했대.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "pmperformance",
      "dm"
    ],
    "keywords": [
      "Benchmark"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "pmperformance": {
    "examples": [
      {
        "where": "",
        "text": "이번 PM 얼마로 잡혔어요?",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "PM 내려가서 샘플 수 좀 늘었어.",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "PM 기준으로 테스트하면 돼요?",
        "by": "익명 선배"
      }
    ],
    "related": [
      "om"
    ],
    "keywords": [
      "Aggregation Risk",
      "Sample",
      "Test of Details"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "dm": {
    "examples": [
      {
        "where": "",
        "text": "이거 DM 밑이라 SUM에는 안 올려도 되죠?",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "이번 DM 얼마예요?",
        "by": "익명 선배"
      }
    ],
    "related": [
      "sum",
      "om",
      "pmperformance"
    ],
    "keywords": [
      "Clearly Trivial",
      "Misstatement"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "sum": {
    "examples": [
      {
        "where": "",
        "text": "이 조정사항 SUM에 올렸어요?",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "클라가 수정 안 한대서 SUM으로 남겨야 할 것 같아요.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "dm",
      "om"
    ],
    "keywords": [
      "Audit Adjustment",
      "Misstatement",
      "Completion"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  },
  "minorpass": {
    "examples": [
      {
        "where": "",
        "text": "이거 원칙상 틀리긴 한데 금액 작아서 minor pass 가능할까요?",
        "by": "익명 선배"
      },
      {
        "where": "",
        "text": "이 정도 차이는 minor pass로 가도 될 것 같아.",
        "by": "익명 선배"
      }
    ],
    "related": [
      "om",
      "dm",
      "sum"
    ],
    "keywords": [
      "Misstatement",
      "Clearly Trivial",
      "Audit Adjustment",
      "질적 중요성"
    ],
    "history": [
      {
        "date": "2026-10-03",
        "by": "익명 선배",
        "what": "용어집에 등록했어요"
      }
    ]
  }
};

const POPULAR = ["geumjo", "assign", "kongjjikgi"];

/* 다른 신입이 이미 요청한 단어 (요청 기능이 서버와 연결되면 채워져요) */
const REQUESTS = [];
