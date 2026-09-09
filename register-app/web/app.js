const rows = document.querySelector('#student-rows');
const template = document.querySelector('#student-row');
const feedback = document.querySelector('#feedback');
const submitButton = document.querySelector('#submit');

function addStudent() {
  const row = template.content.cloneNode(true);
  row.querySelector('.remove').addEventListener('click', (event) => event.target.closest('tr').remove());
  rows.append(row);
}

for (let i = 0; i < 5; i += 1) addStudent();
document.querySelector('#date').value = new Date().toISOString().slice(0, 10);
document.querySelector('#add-student').addEventListener('click', addStudent);

document.querySelector('#register-form').addEventListener('submit', async (event) => {
  event.preventDefault();
  const students = [...rows.querySelectorAll('tr')]
    .map((row) => ({
      name: row.querySelector('.student-name').value,
      status: row.querySelector('.student-status').value,
      time: row.querySelector('.student-time').value,
      notes: row.querySelector('.student-notes').value,
    }))
    .filter((student) => student.name.trim());
  if (!students.length) {
    feedback.textContent = 'Add at least one student name.';
    feedback.className = 'error';
    return;
  }
  submitButton.disabled = true;
  feedback.textContent = 'Preparing Excel register…';
  feedback.className = '';
  try {
    const response = await fetch('/submit', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        className: document.querySelector('#className').value,
        date: document.querySelector('#date').value,
        teacher: document.querySelector('#teacher').value,
        students,
      }),
    });
    if (!response.ok) {
      const result = await response.json();
      throw new Error(result.error || 'Could not export the register.');
    }
    const file = await response.blob();
    const filename = response.headers.get('content-disposition')?.match(/filename="?([^";]+)"?/)?.[1] || 'register.xlsx';
    const link = document.createElement('a');
    link.href = URL.createObjectURL(file);
    link.download = filename;
    link.click();
    URL.revokeObjectURL(link.href);
    feedback.textContent = `Excel register downloaded. Attach it to an email for ala.elkurd@morleycollege.ac.uk.`;
    feedback.className = 'success';
  } catch (error) {
    feedback.textContent = error instanceof TypeError && error.message === 'Failed to fetch'
      ? 'The register service is not reachable. Open http://127.0.0.1:8081 (not the index.html file directly) and leave its launch window open.'
      : error.message;
    feedback.className = 'error';
  } finally {
    submitButton.disabled = false;
  }
});
