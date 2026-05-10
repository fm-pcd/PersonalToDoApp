// ── Todos – Google Apps Script backend ──────────────────────────────────────
// Paste this entire file into your Apps Script project (script.google.com)
// then deploy as a Web App (Execute as: Me, Access: Anyone).
//
// IMPORTANT: change SECRET_TOKEN to a long random string before deploying.
// Deployed URL: https://script.google.com/macros/s/AKfycby-rO7kWuqpz2T-rTJw72h0mC5fOP0YKSyA8lSkEcXCUzJYTSqqL6RZOdo-DezLxb0d/exec

const SECRET_TOKEN = 'florian-todos-xk9m2p4q8r';
const SHEET_NAME   = 'Todos';

// ── Entry point ──────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    if (e.parameter.token !== SECRET_TOKEN) return out({ error: 'Unauthorized' });

    switch (e.parameter.action) {
      case 'list':   return out(listTodos());
      case 'add':    return out(addTodo(e.parameter));
      case 'update': return out(updateTodo(e.parameter));
      case 'delete': return out(deleteTodo(e.parameter));
      default:       return out({ error: 'Unknown action' });
    }
  } catch (err) {
    return out({ error: err.message });
  }
}

function out(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Sheet helpers ─────────────────────────────────────────────────────────────
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET_NAME);
  if (!s) {
    s = ss.insertSheet(SHEET_NAME);
    s.appendRow(['ID', 'Title', 'Notes', 'Source', 'Due Date', 'Done', 'Created At']);
  }
  return s;
}

// ── CRUD ──────────────────────────────────────────────────────────────────────
function listTodos() {
  const rows = getSheet().getDataRange().getValues();
  if (rows.length < 2) return { todos: [] };
  return {
    todos: rows.slice(1).map(r => ({
      id:         String(r[0]),
      title:      String(r[1]),
      notes:      String(r[2]),
      source:     String(r[3]),
      due_date:   String(r[4]),
      done:       String(r[5]),
      created_at: String(r[6])
    }))
  };
}

function addTodo(p) {
  const id = String(Date.now());
  getSheet().appendRow([
    id,
    p.title  || '',
    p.notes  || '',
    p.source || 'manual',
    p.due    || '',
    'FALSE',
    new Date().toISOString()
  ]);
  return { success: true, id };
}

function updateTodo(p) {
  const s    = getSheet();
  const rows = s.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(p.id)) continue;
    const row = i + 1;
    if (p.title !== undefined) s.getRange(row, 2).setValue(p.title);
    if (p.notes !== undefined) s.getRange(row, 3).setValue(p.notes);
    if (p.due   !== undefined) s.getRange(row, 5).setValue(p.due);
    if (p.done  !== undefined) s.getRange(row, 6).setValue(p.done === 'true' ? 'TRUE' : 'FALSE');
    return { success: true };
  }
  return { error: 'Not found' };
}

function deleteTodo(p) {
  const s    = getSheet();
  const rows = s.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(p.id)) {
      s.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: 'Not found' };
}
