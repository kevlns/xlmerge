"""Execute the UI review/navigation logic, including old manifest flags."""
from pathlib import Path
import subprocess
import unittest


class ReviewNavigationTests(unittest.TestCase):
    def test_review_tabs_confirmation_navigation_and_pagination(self):
        html = (Path(__file__).parents[1] / 'xlsx_resolver' / 'merge_ui.html').read_text(encoding='utf-8')
        script = html.split('<script>')[1].split('</script>')[0].split("byId('allLocalBtn').onclick")[0]
        checks = r"""
const assert = require('node:assert/strict');
class Element {
  constructor() { this.children = []; this.dataset = {}; this.classList = {add() {}, remove() {}}; }
  append(...children) { this.children.push(...children); }
  replaceChildren(...children) { this.children = children; }
  setAttribute() {}
  remove() {}
}
const elements = new Map();
const document = {
  getElementById(id) { if (!elements.has(id)) elements.set(id, new Element()); return elements.get(id); },
  createElement() { return new Element(); }, createTextNode(text) { return {textContent: text}; }
};
const sheet = {name: 'One-sided', action: 'merge', needsAttention: false,
  cells: Array.from({length: 501}, (_, i) => ({coordinate: `A${i+1}`, row: i+1,
    col: 1, resolution: i % 2 ? 'ours' : 'theirs', ours: 1, theirs: 2}))};
const hidden = {name: 'Identical', action: 'merge', needsAttention: false,
  cells: [{resolution: 'same', selectable: false}]};
const file = {id: 'a', path: 'a.xlsx', diff: {sheets: [sheet, hidden]}};
const other = {id: 'b', diff: {sheets: [{name: 'Conflict', action: 'conflict_delete_modify'}]}};
model = {files: [file, other]};
initializeDecisions();
renderSheetTabs();
assert.equal(byId('sheetTabs').children.length, 1);
const checkbox = byId('sheetTabs').children[0].children[0];
updateGlobalConfirmationStats();
assert.equal(byId('totalConflictSheetCount').textContent, '2');
checkbox.checked = true;
render = () => {};
checkbox.onchange({stopPropagation() {}});
assert.equal(currentSheetIsConfirmed(), true);
updateGlobalConfirmationStats();
assert.equal(byId('confirmedSheetCount').textContent, '1');
renderActionAvailability();
assert.equal(byId('allLocalBtn').disabled, true);
checkbox.checked = false;
checkbox.onchange({stopPropagation() {}});
assert.equal(currentSheetIsConfirmed(), false);
renderConflictNavigation();
assert.equal(byId('nextConflictBtn').disabled, false);
assert.equal(byId('conflictNavCount').textContent, '0 / 501');
renderLegacyDiffTable(byId('tableContainer'), file, sheet);
const firstCell = byId('tableContainer').children[0].children[1].children[0].children[1];
assert.equal(firstCell.dataset.conflictId, 'A1');
let animation;
const requestAnimationFrame = callback => { animation = callback; };
renderTable = () => {};
locateConflictCell(-1);
assert.equal(currentPage, 1);
assert.equal(byId('conflictNavCount').textContent, '501 / 501');
locateConflictCell(1);
assert.equal(currentPage, 0);
assert.equal(byId('conflictNavCount').textContent, '1 / 501');
// Locked structural additions/removals and automatic repairs are not targets.
const structural = {name: 'Structure', action: 'merge', needsAttention: false,
  allRows: [{id: 'r', state: 'common'}, {id: 'locked', state: 'added', locked: true}],
  columns: [{id: 'c', state: 'common'}],
  cells: ['ours', 'theirs', 'conflict', 'same', 'replacement'].map((resolution, i) =>
    ({id: String(i), rowId: 'r', colId: 'c', resolution, changed: true,
      selectable: resolution !== 'same', replacement: resolution === 'replacement'}))};
structural.cells.push({id: 'locked', rowId: 'locked', colId: 'c', resolution: 'theirs', changed: true});
assert.deepEqual(navigableConflictCells(file, structural).map(cell => cell.id), ['0', '1', '2']);
sheetDecision(file, structural).rows.r = 'delete';
// Changing a decision never hides the Sheet's confirmation requirement.
assert.equal(sheetNeedsConfirmation(structural), true);
"""
        result = subprocess.run(['node'], input=script + checks, text=True, encoding='utf-8', capture_output=True)
        self.assertEqual(result.returncode, 0, result.stderr)


if __name__ == '__main__':
    unittest.main()
