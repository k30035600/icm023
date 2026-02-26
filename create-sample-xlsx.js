const fs = require('fs');
const path = require('path');
// CDN 버전과 동일한 SheetJS 사용 (npm xlsx)
let XLSX;
try {
  XLSX = require('xlsx');
} catch (e) {
  console.error('xlsx 패키지가 필요합니다. 실행: npm install xlsx');
  process.exit(1);
}

const leftData = [
  ['구분', '고객명', '연락처', '담당자', '비고'],
  ['A', '김고객', '02-1234-5678', '홍길동', 'VIP'],
  ['B', '이고객', '02-2345-6789', '김담당', ''],
  ['C', '박고객', '031-111-2222', '이담당', '재연락'],
];

const rightData = [
  ['일자', '접촉내용', '결과', '다음일정'],
  ['2025-02-20', '전화 상담', '긍정적', '2025-03-01'],
  ['2025-02-22', '방문 상담', '검토 중', '2025-02-28'],
  ['2025-02-25', '이메일 발송', '대기', ''],
];

function writeSheet(filename, rows) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  const outPath = path.join(__dirname, filename);
  XLSX.writeFile(wb, outPath);
  console.log('생성됨:', outPath);
}

writeSheet('좌.xlsx', leftData);
writeSheet('우.xlsx', rightData);
