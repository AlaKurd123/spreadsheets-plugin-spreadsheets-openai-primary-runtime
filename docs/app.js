const rows = document.querySelector('#student-rows');
const template = document.querySelector('#student-row');
const feedback = document.querySelector('#feedback');
const submitButton = document.querySelector('#submit');

function addStudent() {
  const row = template.content.cloneNode(true);
  row.querySelector('.remove').addEventListener('click', (event) => event.target.closest('tr').remove());
  rows.append(row);
}

function escapeXml(value) {
  return String(value ?? '').replace(/[<>&"']/g, (character) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;', "'": '&apos;' }[character]));
}

function columnName(index) {
  return String.fromCharCode(65 + index);
}

function xmlCell(column, row, value, style = 0) {
  return `<c r="${column}${row}" t="inlineStr"${style ? ` s="${style}"` : ''}><is><t>${escapeXml(value)}</t></is></c>`;
}

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function dosTime(date) {
  return ((date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2)) & 0xffff;
}

function dosDate(date) {
  return (((date.getFullYear() - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate()) & 0xffff;
}

function u16(value) { return Uint8Array.of(value & 255, (value >>> 8) & 255); }
function u32(value) { return Uint8Array.of(value & 255, (value >>> 8) & 255, (value >>> 16) & 255, (value >>> 24) & 255); }
function join(parts) {
  const size = parts.reduce((total, part) => total + part.length, 0);
  const result = new Uint8Array(size);
  let offset = 0;
  for (const part of parts) { result.set(part, offset); offset += part.length; }
  return result;
}

function zip(files) {
  const encoder = new TextEncoder();
  const now = new Date();
  let offset = 0;
  const localParts = [];
  const centralParts = [];
  for (const [name, content] of files) {
    const nameBytes = encoder.encode(name);
    const data = encoder.encode(content);
    const crc = crc32(data);
    const local = join([u32(0x04034b50), u16(20), u16(0), u16(0), u16(dosTime(now)), u16(dosDate(now)), u32(crc), u32(data.length), u32(data.length), u16(nameBytes.length), u16(0), nameBytes, data]);
    localParts.push(local);
    centralParts.push(join([u32(0x02014b50), u16(20), u16(20), u16(0), u16(0), u16(dosTime(now)), u16(dosDate(now)), u32(crc), u32(data.length), u32(data.length), u16(nameBytes.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(offset), nameBytes]));
    offset += local.length;
  }
  const central = join(centralParts);
  return join([...localParts, central, u32(0x06054b50), u16(0), u16(0), u16(files.length), u16(files.length), u32(central.length), u32(offset), u16(0)]);
}

function createWorkbook(data) {
  const detailRows = [['Class', data.className], ['Date', data.date], ['Teacher', data.teacher], ['Exported', new Date().toLocaleString('en-GB')]];
  const headers = ['Student name', 'Status', 'Time', 'Notes'];
  let sheetRows = `<row r="1">${xmlCell('A', 1, 'Temporary Register', 1)}</row>`;
  detailRows.forEach(([label, value], index) => { const row = index + 3; sheetRows += `<row r="${row}">${xmlCell('A', row, label, 2)}${xmlCell('B', row, value)}</row>`; });
  sheetRows += `<row r="8">${headers.map((header, index) => xmlCell(columnName(index), 8, header, 3)).join('')}</row>`;
  data.students.forEach((student, index) => {
    const row = index + 9;
    const statusStyle = student.status === 'Absent' ? 4 : student.status === 'Late' ? 5 : 0;
    sheetRows += `<row r="${row}">${xmlCell('A', row, student.name)}${xmlCell('B', row, student.status, statusStyle)}${xmlCell('C', row, student.time)}${xmlCell('D', row, student.notes)}</row>`;
  });
  const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="8" topLeftCell="A9" state="frozen"/></sheetView></sheetViews><cols><col min="1" max="1" width="28" customWidth="1"/><col min="2" max="2" width="14" customWidth="1"/><col min="3" max="3" width="14" customWidth="1"/><col min="4" max="4" width="45" customWidth="1"/></cols><sheetData>${sheetRows}</sheetData><mergeCells count="1"><mergeCell ref="A1:D1"/></mergeCells></worksheet>`;
  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Aptos"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font></fonts><fills count="6"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF5B9BD5"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFCE4D6"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="1"><alignment/></xf><xf numFmtId="0" fontId="1" fillId="3" borderId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="4" borderId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="5" borderId="0" applyFill="1"/></cellXfs></styleSheet>`;
  return zip([
    ['[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'],
    ['_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
    ['xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Register" sheetId="1" r:id="rId1"/></sheets></workbook>'],
    ['xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'],
    ['xl/styles.xml', styles],
    ['xl/worksheets/sheet1.xml', sheet],
  ]);
}

function downloadWorkbook(data) {
  const file = new Blob([createWorkbook(data)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const filename = `${data.className.trim().replace(/[^A-Za-z0-9_-]+/g, '-') || 'register'}_${data.date}.xlsx`;
  const link = document.createElement('a');
  link.href = URL.createObjectURL(file);
  link.download = filename;
  link.click();
  URL.revokeObjectURL(link.href);
}

for (let i = 0; i < 5; i += 1) addStudent();
document.querySelector('#date').value = new Date().toISOString().slice(0, 10);
document.querySelector('#add-student').addEventListener('click', addStudent);

document.querySelector('#register-form').addEventListener('submit', (event) => {
  event.preventDefault();
  const students = [...rows.querySelectorAll('tr')].map((row) => ({ name: row.querySelector('.student-name').value.trim(), status: row.querySelector('.student-status').value, time: row.querySelector('.student-time').value, notes: row.querySelector('.student-notes').value.trim() })).filter((student) => student.name);
  if (!students.length) {
    feedback.textContent = 'Add at least one student name.';
    feedback.className = 'error';
    return;
  }
  const data = { className: document.querySelector('#className').value.trim(), date: document.querySelector('#date').value, teacher: document.querySelector('#teacher').value.trim(), students };
  downloadWorkbook(data);
  feedback.textContent = 'Excel register downloaded. Attach it to an email for ala.elkurd@morleycollege.ac.uk.';
  feedback.className = 'success';
});
