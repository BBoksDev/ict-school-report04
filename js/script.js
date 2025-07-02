document.addEventListener('DOMContentLoaded', () => {
  const table = document.getElementById('spreadsheet');
  const tbodyRows = table.querySelectorAll('tbody tr');

  tbodyRows.forEach((tr, rowIndex) => {
    const cells = tr.querySelectorAll('td');

    cells.forEach((td, colIndex) => {
      // 셀 포커스 시 highlight
      td.addEventListener('focus', () => {
        highlight(rowIndex + 1, colIndex);
      });

      // 셀 키보드 이동
      td.addEventListener('keydown', e => {
        const rowCount = tbodyRows.length;
        const colCount = cells.length;
        let newRow = rowIndex;
        let newCol = colIndex;

        if (e.key === 'Tab') {
          e.preventDefault();
          newCol += e.shiftKey ? -1 : 1;
        } else if (e.key === 'Enter') {
          e.preventDefault();
          newRow += e.shiftKey ? -1 : 1;
        } else {
          switch (e.key) {
            case 'ArrowRight': newCol++; e.preventDefault(); break;
            case 'ArrowLeft':  newCol--; e.preventDefault(); break;
            case 'ArrowDown':  newRow++; e.preventDefault(); break;
            case 'ArrowUp':    newRow--; e.preventDefault(); break;
            default: return;
          }
        }

        // 유효한 셀로 이동
        if (newRow >= 0 && newRow < rowCount) {
          const nextRow = tbodyRows[newRow];
          const nextCell = nextRow.querySelectorAll('td')[newCol];
          if (nextCell) nextCell.focus();
        }
      });

      // 포커스 해제 시 highlight 제거
      td.addEventListener('focusout', removeHighlight);
    });
  });
});

// highlight 주기
function highlight(row, col) {
  const table = document.getElementById('spreadsheet');
  const ths = table.querySelectorAll('th');

  ths.forEach(th => th.classList.remove('highlight'));

  const columnHeader = table.querySelector(`thead tr th:nth-child(${col + 2})`);
  const rowHeader = table.querySelector(`tbody tr:nth-child(${row}) th`);

  if (columnHeader) columnHeader.classList.add('highlight');
  if (rowHeader) rowHeader.classList.add('highlight');

  const cellIdSpan = document.getElementById('cellId');
  if (columnHeader && rowHeader) {
    cellIdSpan.innerText = `${columnHeader.innerText}${rowHeader.innerText}`;
  }
}

// highlight 제거
function removeHighlight() {
  const cellIdSpan = document.getElementById('cellId');
  cellIdSpan.innerText = '';

  const table = document.getElementById('spreadsheet');
  const ths = table.querySelectorAll('th');
  ths.forEach(th => th.classList.remove('highlight'));
}

// Excel로 추출
function exportToExcel() {
  const table = document.getElementById('spreadsheet');
  const wb = XLSX.utils.book_new();
  const rows = [];

  table.querySelectorAll('tbody tr').forEach(tr => {
    const rowData = [];
    tr.querySelectorAll('td').forEach(td => {
      rowData.push(td.innerText || '');
    });
    rows.push(rowData);
  });

  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  XLSX.writeFile(wb, 'spreadsheet.xlsx');
}
