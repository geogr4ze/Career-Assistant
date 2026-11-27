const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');

const source = fs.readFileSync(path.join(__dirname, 'EmailMonitor.gs'), 'utf8');
const ctx = vm.createContext({
  Utilities: { formatDate: () => '2026-11-26' }
});

/**
 * Evaluate isolated Apps Script source or a test expression without provider access.
 * @param {string} expression - JavaScript source.
 * @returns {*} Evaluation result.
 * @throws {Error} If evaluation fails.
 */
function evaluate(expression) {
  return vm.runInContext(expression, ctx);
}

evaluate(source);
evaluate(`
  const app = Array(24).fill('');
  app[0] = 'A1';
  app[1] = 'NVIDIA';
  app[2] = 'Firmware <intern>';
  app[11] = '2026-11-26';
  app[13] = 'confirmation';
  const response = Array(11).fill('');
  response[0] = 'R1';
  response[1] = 'A1';
  response[2] = '2026-11-26 10:00';
  response[4] = 'Interview invitation';
  response[6] = '<script>alert(1)</script>';
  response[8] = 'Choose time';
  response[9] = '2026-12-01';
  const events = emEvents_([app], [response, response]);
  const message = emMessage_(
    events.find(event => event.id === 'RESPONSE:R1'),
    'https://docs.google.com/spreadsheets/d/EXAMPLE/edit'
  );
  const mockBook = {
    getSheetByName: () => ({
      getDataRange: () => ({
        getDisplayValues: () => [['Date'], ['2026-11-26'], ['invalid']],
        getValues: () => [['Date'], [new Date('2026-11-26')], ['invalid']]
      })
    })
  };
  const dateRows = emRows_(mockBook, 'Test', {0: 'Date'}, [0]);
`);

assert.equal(evaluate('events.length'), 2);
assert.equal(evaluate('emEvents_([app], [["R2", "UNKNOWN"]]).length'), 1);
assert.equal(evaluate('emEvents_([app.slice(0, 13)], []).length'), 0);
assert.equal(evaluate('emPriority_("NVIDIA")'), true);
assert.equal(evaluate('emPriority_("Unknown")'), false);
assert.match(evaluate('message.subject'), /^PRIORITY/);
assert.equal(evaluate('message.html.includes("<script>")'), false);
assert.match(evaluate('message.html'), /&lt;script&gt;/);
assert.match(evaluate('message.text'), /Choose time/);
assert.match(evaluate('message.html'), /View application dashboard/);
assert.equal(evaluate('dateRows[0][0]'), '2026-11-26');
assert.equal(evaluate('dateRows[1][0]'), '');
assert.equal(evaluate('emDay_(new Date())'), 'PT:2026-11-26');
console.log('13 email monitor checks passed. No external calls made.');
