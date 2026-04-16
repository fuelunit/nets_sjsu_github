// ─────────────────────────────────────────────────────────────
// CLOUDATHON TEAM ASSIGNMENT
// ─────────────────────────────────────────────────────────────
// Author: Linh (Maria) Truong <email@hidden>
// Contributors: Yipeng Liu <yipeng.liu@sjsu.edu>
//               Rishik R. Dammannagari <email@hidden>
// Date Created: 04/09/2026
// Date Updated: 04/15/2026
// WORKFLOW (every time):
//   1. Clear all rows below the header in Registrants and Teams
//   2. Run importFromFormResponses()
//   3. Run runAssignment()
//   4. Check MailMerge for anything flagged, fix manually
//
// REGISTRANTS columns (never reorder):
//   A email  B first  C last  D competition  E skill  F status
//   G team_id  H leader_email  I open_to_merge  J registered_at  K (unused)  L flags
//
// TEAMS columns:
//   A team_id  B competition  C member_emails  D team_type
// ─────────────────────────────────────────────────────────────

const C = {
  EMAIL: 0, FIRST: 1, LAST: 2, COMP: 3, SKILL: 4,
  STATUS: 5, TEAM: 6, LEADER: 7, OPEN: 8, TIME: 9,
  OFFER: 10, FLAGS: 11
};

// Update these if form question order ever changes
const F = {
  TIMESTAMP: 0, FIRST: 2, LAST: 4, EDU_EMAIL: 5,
  COMP: 8, SKILL: 9, PLACEMENT: 10, LEADER: 11, OPEN: 12
};

const SHEET = name => SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);

function allRows() {
  const s = SHEET('Registrants');
  const last = s.getLastRow();
  return last < 2 ? [] : s.getRange(2, 1, last - 1, 12).getValues();
}

function setCell(rowIdx, colIdx, value) {
  SHEET('Registrants').getRange(rowIdx + 2, colIdx + 1).setValue(value);
}

function norm(v)        { return String(v ?? '').trim().toLowerCase(); }
function compShort(raw) {
  const s = norm(raw);
  if (s.includes('cyber'))   return 'Cybersecurity';
  if (s.includes('cloud'))   return 'Cloud';
  if (s.includes('network')) return 'Network';
  return '';
}

function levenshtein(a, b) {
  const m = a.length, n = b.length;
  const d = Array.from({length: m + 1}, (_, i) => [i]);
  for (let j = 0; j <= n; j++) d[0][j] = j;
  for (let i = 1; i <= m; i++)
    for (let j = 1; j <= n; j++)
      d[i][j] = a[i-1] === b[j-1] ? d[i-1][j-1]
        : 1 + Math.min(d[i-1][j], d[i][j-1], d[i-1][j-1]);
  return d[m][n];
}

// ─────────────────────────────────────────────────────────────
// importFromFormResponses
// ─────────────────────────────────────────────────────────────
function importFromFormResponses() {
  const formSheet = SHEET('Form Responses 1');
  const regSheet  = SHEET('Registrants');

  if (!formSheet) {
    SpreadsheetApp.getUi().alert('Tab "Form Responses 1" not found.'); return;
  }
  if (regSheet.getLastRow() > 1) {
    SpreadsheetApp.getUi().alert('Registrants already has data. Clear rows 2 onwards first.'); return;
  }

  const rows = formSheet.getDataRange().getValues().slice(1);
  if (rows.length === 0) {
    SpreadsheetApp.getUi().alert('No responses found.'); return;
  }

  // Group by edu email — latest submission wins per person
  const byEdu = {};
  rows.forEach(v => {
    const email = norm(v[F.EDU_EMAIL]);
    if (email) (byEdu[email] = byEdu[email] || []).push(v);
  });

  // Find resubmits: same person, different comp or leader across submissions
  const resubmitEmails = new Set();
  Object.entries(byEdu).forEach(([email, subs]) => {
    if (subs.length < 2) return;
    const comps   = [...new Set(subs.map(v => compShort(v[F.COMP])))];
    const leaders = [...new Set(subs.map(v =>
      norm(v[F.PLACEMENT]).startsWith('no') ? norm(v[F.LEADER] || '') : '__random__'
    ))];
    if (comps.length > 1 || leaders.length > 1) resubmitEmails.add(email);
  });

  const allEduEmails = Object.keys(byEdu);

  let count = 0;
  Object.entries(byEdu).forEach(([email, subs]) => {
    const v    = subs[subs.length - 1]; // latest submission
    const comp = compShort(v[F.COMP]);
    if (!comp) return;

    const skill     = (v[F.SKILL]  || '').trim();
    const isPreform = norm(v[F.PLACEMENT]).startsWith('no');
    const firstName = (v[F.FIRST]  || '').trim();
    const lastName  = (v[F.LAST]   || '').trim();
    const time      = v[F.TIMESTAMP] || new Date();
    const openMerge = norm(v[F.OPEN]).includes('yes');

    let leader = isPreform ? norm(v[F.LEADER] || '') : '';
    let flag   = resubmitEmails.has(email)
      ? `RESUBMIT: ${subs.length} submissions, keeping latest (${comp}${leader ? ', leader=' + leader : ''})`
      : '';

    // Typo correction: if leader email not found but close to a known email,
    // auto-correct so the student still joins the intended group
    if (leader && !byEdu[leader]) {
      const match = allEduEmails.find(e => e !== email && levenshtein(e, leader) <= 2);
      if (match) {
        flag = (flag ? flag + ' | ' : '') + `TYPO: leader "${leader}" auto-corrected to "${match}"`;
        leader = match;
      }
    }

    count++;
    regSheet.appendRow([
      email, firstName, lastName, comp, skill,
      'Confirmed', '', leader, openMerge, time, '', flag
    ]);
  });

  SpreadsheetApp.getUi().alert(`Imported ${count} rows.\nNow run runAssignment().`);
}

// ─────────────────────────────────────────────────────────────
// runAssignment
// ─────────────────────────────────────────────────────────────
function runAssignment() {
  const teamSheet = SHEET('Teams');
  if (teamSheet.getLastRow() > 1) {
    SpreadsheetApp.getUi().alert('Teams already has data. Clear it first.'); return;
  }

  const rows = allRows().filter(r => r[C.STATUS] === 'Confirmed');
  if (rows.length === 0) {
    SpreadsheetApp.getUi().alert('No Confirmed students found.'); return;
  }

  // ── Build group map ───────────────────────────────────────
  const groupMap = {};
  rows.filter(r => r[C.LEADER] !== '').forEach(r => {
    const key = norm(r[C.LEADER]);
    (groupMap[key] = groupMap[key] || []).push(r);
  });

  // ── Detect mutual leaders (A listed B, B listed A) ────────
  const leaderOf      = {};
  rows.forEach(r => { if (r[C.LEADER]) leaderOf[norm(r[C.EMAIL])] = norm(r[C.LEADER]); });

  const mutualFlagged = new Set();
  const mutualPairs   = [];
  Object.entries(leaderOf).forEach(([email, leader]) => {
    if (leaderOf[leader] !== email || email === leader) return;
    const pairKey = [email, leader].sort().join('|');
    if (mutualFlagged.has(pairKey)) return;
    mutualFlagged.add(pairKey);
    mutualFlagged.add(email);
    mutualFlagged.add(leader);

    const rowA = rows.find(r => norm(r[C.EMAIL]) === email);
    const rowB = rows.find(r => norm(r[C.EMAIL]) === leader);
    [email, leader].forEach(e => {
      const other  = e === email ? leader : email;
      const i = rows.findIndex(r => norm(r[C.EMAIL]) === e);
      if (i !== -1) setCell(i, C.FLAGS,
        `MUTUAL_LEADER: ${e} and ${other} listed each other — no valid group`);
    });
    const leaderComp = rowA ? rowA[C.COMP] : (rowB ? rowB[C.COMP] : null);
    if (leaderComp && rowA && rowB) mutualPairs.push({ leaderComp, members: [rowA, rowB] });
  });

  // ── Classify each group ───────────────────────────────────
  const fixedByComp   = { Cybersecurity: [], Cloud: [], Network: [] };
  const poolByComp    = { Cybersecurity: [], Cloud: [], Network: [] };
  const flaggedGroups = [];

  Object.entries(groupMap).forEach(([leaderEmail, members]) => {
    const leaderRow  = rows.find(r => norm(r[C.EMAIL]) === leaderEmail);
    const leaderComp = leaderRow ? leaderRow[C.COMP] : '';

    // Shared helper: write flag on all members and send group to FLAGGED-
    const flagAll = (msg, comp) => {
      members.forEach(r => {
        const i = rows.findIndex(x => norm(x[C.EMAIL]) === norm(r[C.EMAIL]));
        if (i === -1) return;
        const existing = rows[i][C.FLAGS];
        setCell(i, C.FLAGS, existing ? `${existing} | ${msg}` : msg);
      });
      flaggedGroups.push({ leaderComp: comp || leaderComp, members });
    };

    if (!leaderComp)                                          { flagAll(`LEADER_NOT_FOUND: ${leaderEmail} has not registered`, members[0][C.COMP]); return; }
    if (mutualFlagged.has(leaderEmail))                       { return; }
    if (members.some(r => r[C.COMP] !== leaderComp))          { flagAll(`CROSS_COMP: group spans multiple competitions (leader ${leaderEmail} is in ${leaderComp})`); return; }
    if (leaderRow[C.FLAGS] && leaderRow[C.FLAGS].includes('RESUBMIT')) { flagAll(`RESUBMIT_LEADER: ${leaderEmail} changed their registration`); return; }
    if (members.length > 5)                                   { flagAll(`OVERSIZED: ${members.length} members (max 5)`); return; }

    // Clean group — Bucket A (≥4 or closed): fixed; Bucket B (<4, open): dissolve into pool
    const anyOpen = members.some(r => r[C.OPEN] === true);
    if (members.length >= 4 || !anyOpen) {
      fixedByComp[leaderComp].push(members);
    } else {
      poolByComp[leaderComp].push(...members);
    }
  });

  mutualPairs.forEach(p => flaggedGroups.push(p));

  // ── Build pool teams per competition ─────────────────────
  const allTeams = [];
  ['Cybersecurity', 'Cloud', 'Network'].forEach(comp => {
    const prefix = comp === 'Cybersecurity' ? 'CYB' : comp === 'Cloud' ? 'CLD' : 'NET';
    let n_ = 1;

    const pool = [
      ...rows.filter(r => r[C.COMP] === comp && r[C.LEADER] === ''),
      ...poolByComp[comp]
    ];
    pool.sort((a, b) => ({'Advanced':0,'Intermediate':1,'Beginner':2}[a[C.SKILL]] ?? 3)
                       - ({'Advanced':0,'Intermediate':1,'Beginner':2}[b[C.SKILL]] ?? 3));

    const nTeams    = pool.length <= 5 ? (pool.length > 0 ? 1 : 0) : Math.ceil(pool.length / 4);
    const poolTeams = Array.from({ length: nTeams }, () => []);
    pool.forEach((s, i) => poolTeams[i % nTeams].push(s));

    for (let i = poolTeams.length - 1; i >= 0; i--) {
      if (poolTeams[i].length <= 2 && poolTeams.length > 1) {
        const target = poolTeams.filter((_, j) => j !== i).sort((a, b) => a.length - b.length)[0];
        if (target && target.length + poolTeams[i].length <= 5) {
          target.push(...poolTeams[i]);
          poolTeams.splice(i, 1);
        }
      }
    }

    [...fixedByComp[comp], ...poolTeams].forEach(members =>
      allTeams.push({ id: `${prefix}-${String(n_++).padStart(2, '0')}`, comp, members })
    );
    flaggedGroups.filter(g => g.leaderComp === comp).forEach(g =>
      allTeams.push({ id: `FLAGGED-${prefix}`, comp, members: g.members, isFlagged: true })
    );
  });

  // ── Flagged emails win: strip from regular teams ──────────
  const flaggedEmails = new Set(
    allTeams.filter(t => t.isFlagged).flatMap(t => t.members.map(r => norm(r[C.EMAIL])))
  );
  allTeams.forEach(t => {
    if (!t.isFlagged) t.members = t.members.filter(r => !flaggedEmails.has(norm(r[C.EMAIL])));
  });
  const teamsToWrite = allTeams.filter(t => t.members.length > 0);

  // ── Write Teams sheet ─────────────────────────────────────
  teamsToWrite.forEach(t => {
    const preCount = t.members.filter(r => r[C.LEADER] !== '').length;
    const teamType = t.isFlagged ? 'Flagged'
                   : preCount === t.members.length ? 'Pre-set'
                   : preCount === 0               ? 'Randomized'
                                                  : 'Mixed';
    teamSheet.appendRow([t.id, t.comp, t.members.map(r => r[C.EMAIL]).join(', '), teamType]);
  });

  // ── Write team_id back to Registrants ────────────────────
  const emailToTeam = {};
  teamsToWrite.forEach(t => t.members.forEach(r => emailToTeam[norm(r[C.EMAIL])] = t.id));
  allRows().forEach((row, i) => {
    const tid = emailToTeam[norm(row[C.EMAIL])];
    if (tid) setCell(i, C.TEAM, tid);
  });

  const flagCount = teamsToWrite.filter(t => t.isFlagged).length;
  SpreadsheetApp.getUi().alert(
    `Done. ${teamsToWrite.length} teams created (${flagCount} flagged).\nCheck MailMerge for anything needing review.`
  );
}