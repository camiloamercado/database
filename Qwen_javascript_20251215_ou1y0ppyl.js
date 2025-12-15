document.addEventListener('DOMContentLoaded', () => {
  const fileAInput = document.getElementById('fileA');
  const fileBInput = document.getElementById('fileB');
  const compareBtn = document.getElementById('compareBtn');
  const resultsDiv = document.getElementById('results');

  // Enable button only when both files selected
  [fileAInput, fileBInput].forEach(input => {
    input.addEventListener('change', () => {
      compareBtn.disabled = !(fileAInput.files.length && fileBInput.files.length);
    });
  });

  compareBtn.addEventListener('click', async () => {
    try {
      resultsDiv.innerHTML = '<p>⏳ Processing...</p>';

      const fileA = await readFile(fileAInput.files[0]);
      const fileB = await readFile(fileBInput.files[0]);

      const dataA = parseFileA(fileA);
      const dataB = parseFileB(fileB);

      const comparison = compareData(dataA, dataB);

      renderResults(comparison);
    } catch (err) {
      resultsDiv.innerHTML = `<p style="color:red">❌ Error: ${err.message}</p>`;
      console.error(err);
    }
  });

  async function readFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          resolve(workbook);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = () => reject(new Error('Failed to read file'));
      reader.readAsArrayBuffer(file);
    });
  }

  function parseFileA(wb) {
    // First sheet
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

    // Get header row
    const headers = data[0].map(h => String(h).trim().replace(/\n/g, ' '));
    const map = {};
    headers.forEach((h, i) => {
      if (h.includes('Project Code') || h.includes('Agreement')) map.code = i;
      if (h.includes('Pledge') && h.includes('USD')) map.pledge = i;
      if (h.includes('End Date')) map.endDate = i;
      if (h === 'Status') map.status = i;
    });

    const rows = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const code = String(row[map.code] || '').trim();
      if (!code) continue;

      const pledge = parseFloat(row[map.pledge]) * 1000; // → USD
      rows[code] = {
        pledge: isNaN(pledge) ? null : pledge,
        endDate: row[map.endDate],
        status: String(row[map.status] || '').trim()
      };
    }
    return rows;
  }

  function parseFileB(wb) {
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

    const headers = data[0].map(h => String(h).trim());
    const map = {};
    headers.forEach((h, i) => {
      if (h === 'Project Code') map.code = i;
      if (h === 'Original Contract Amount - F') map.pledge = i;
      if (h === 'Date To') map.endDate = i;
      if (h.includes('status') && h.includes('Contract')) map.status = i;
    });

    const rows = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const code = String(row[map.code] || '').trim();
      if (!code) continue;

      const pledge = parseFloat(row[map.pledge]);
      rows[code] = {
        pledge: isNaN(pledge) ? null : pledge,
        endDate: row[map.endDate],
        status: String(row[map.status] || '').trim()
      };
    }
    return rows;
  }

  function compareData(a, b) {
    const allCodes = [...new Set([...Object.keys(a), ...Object.keys(b)])];
    const results = [];

    allCodes.forEach(code => {
      const recA = a[code] || {};
      const recB = b[code] || {};

      const pledgeDiff = recB.pledge - recA.pledge;
      const dateA = recA.endDate ? new Date(recA.endDate) : null;
      const dateB = recB.endDate ? new Date(recB.endDate) : null;
      const dateDiff = dateB && dateA ? dateB - dateA : 0;
      const statusChanged = (recA.status || '') !== (recB.status || '');

      if (pledgeDiff !== 0 || dateDiff !== 0 || statusChanged) {
        results.push({
          code,
          pledgeA: recA.pledge,
          pledgeB: recB.pledge,
          pledgeDiff,
          pledgeChange: pledgeDiff > 0 ? 'Increased' : pledgeDiff < 0 ? 'Decreased' : 'No Change',
          dateA: recA.endDate,
          dateB: recB.endDate,
          dateDiffDays: dateDiff ? Math.round(dateDiff / (1000 * 60 * 60 * 24)) : 0,
          dateChange: dateDiff > 0 ? 'Extended' : dateDiff < 0 ? 'Shortened' : 'No Change',
          statusA: recA.status,
          statusB: recB.status,
          statusChanged
        });
      }
    });

    return results;
  }

  function renderResults(comparison) {
    if (comparison.length === 0) {
      resultsDiv.innerHTML = '<p>✅ No changes detected.</p>';
      return;
    }

    let html = `<h2>📊 ${comparison.length} Projects with Changes</h2>
    <table>
      <thead>
        <tr>
          <th>Project Code</th>
          <th>Pledge Change</th>
          <th>End Date Change</th>
          <th>Status Change</th>
        </tr>
      </thead>
      <tbody>`;

    comparison.forEach(row => {
      const pledgeClass = row.pledgeDiff > 0 ? 'increased' : row.pledgeDiff < 0 ? 'decreased' : '';
      const dateClass = row.dateDiffDays > 0 ? 'extended' : row.dateDiffDays < 0 ? 'shortened' : '';
      const statusClass = row.statusChanged ? 'status-change' : '';

      html += `
        <tr class="${statusClass}">
          <td><strong>${row.code}</strong></td>
          <td class="${pledgeClass}">${row.pledgeChange} ($${Math.abs(row.pledgeDiff).toLocaleString()} USD)</td>
          <td class="${dateClass}">${row.dateChange} (${row.dateDiffDays > 0 ? '+' : ''}${row.dateDiffDays} days)</td>
          <td>${row.statusChanged ? `✅ <s>${row.statusA}</s> → ${row.statusB}` : 'No Change'}</td>
        </tr>`;
    });

    html += `</tbody></table>`;
    resultsDiv.innerHTML = html;
  }
});