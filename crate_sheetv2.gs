/**
 * Crate — vinyl lookup for Kevin's "Vinyl Playlist Building" sheet
 * Discogs (tracklist, styles, label) + GetSongBPM (BPM, key)
 *
 * Setup: in the sheet, Extensions → Apps Script → replace everything with this file → Save
 *        → reload the sheet. A "Crate" menu appears. Run "Set API keys…" once.
 *
 * Writes into your existing tab (the one with Artist / Release / Track / … / Mixes well with).
 * Only fills Artist, Release, Track, Track Artist, Track Name, Genre, Tags, BPM, Key, Camelot,
 * plus a few reference columns it adds at the far right. Mood, Favourite, Playlist and
 * "Mixes well with" are always left for you.
 */

// Your columns (matched by name, so order doesn't matter)
const F = {
  artist: 'Artist', release: 'Release', pos: 'Track', trackArtist: 'Track Artist',
  title: 'Track Name', genre: 'Genre', tags: 'Tags', bpm: 'BPM', key: 'Key', camelot: 'Camelot'
};
// Reference columns added to the right of your existing ones if they're not there yet
const EXTRA = { label: 'Label', catno: 'Cat #', year: 'Year', url: 'Discogs URL', source: 'BPM/Key source' };

const FLAG_COLOUR = '#fde2c8';
const RETRY_LIMIT = 120;   // rows per "Fill missing BPM/key" run (keeps under Apps Script's 6-min limit)

// ---------- Menu ----------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Crate')
    .addItem('Add release…', 'addRelease')
    .addItem('Fill missing BPM/key', 'fillMissing')
    .addItem('Tidy keys + fill Camelot', 'tidyKeys')
    .addItem('Split genres into tags', 'splitGenres')
    .addSeparator()
    .addItem('Set API keys…', 'setKeys')
    .addToUi();
}

function setKeys() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getUserProperties();

  const d = ui.prompt('Discogs personal access token',
    'Discogs → Settings → Developers → Generate new token.\nLeave blank to keep the current one.',
    ui.ButtonSet.OK_CANCEL);
  if (d.getSelectedButton() !== ui.Button.OK) return;
  if (d.getResponseText().trim()) props.setProperty('DISCOGS_TOKEN', d.getResponseText().trim());

  const g = ui.prompt('GetSongBPM API key',
    'From getsongbpm.com/api.\nLeave blank to keep the current one (or skip for now).',
    ui.ButtonSet.OK_CANCEL);
  if (g.getSelectedButton() === ui.Button.OK && g.getResponseText().trim()) {
    props.setProperty('GETSONGBPM_KEY', g.getResponseText().trim());
  }
  ui.alert('Keys saved.');
}

// ---------- Add a release ----------

function addRelease() {
  const ui = SpreadsheetApp.getUi();
  const token = prop_('DISCOGS_TOKEN');
  if (!token) { ui.alert('Add your Discogs token first: Crate → Set API keys…'); return; }

  const r = ui.prompt('Add release',
    'Paste a Discogs release URL (most accurate), or type:  Artist - Release title',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const input = r.getResponseText().trim();
  if (!input) return;

  let releaseId = parseReleaseId_(input);
  if (!releaseId) {
    const candidates = searchDiscogs_(input, token);
    if (!candidates.length) {
      ui.alert('No Discogs match for "' + input + '". Try pasting the release URL instead.');
      return;
    }
    releaseId = pickCandidate_(candidates, ui);
    if (!releaseId) return;
  }

  const rel = discogs_('/releases/' + releaseId, token);
  const sheet = findSheet_();
  const cols = columns_(sheet);
  const width = sheet.getLastColumn();
  const seen = existingTracks_(sheet, cols);
  const bpmKey = prop_('GETSONGBPM_KEY');

  const relArtist = artistNames_(rel.artists);
  const lab = (rel.labels && rel.labels[0]) || {};
  const g = genreAndTags_((rel.styles && rel.styles.length ? rel.styles : (rel.genres || [])).join(', '));
  const tracks = (rel.tracklist || []).filter(t => t.type_ === 'track');

  const rows = [];
  let skipped = 0;
  tracks.forEach(t => {
    if (seen.has(trackId_(rel.title, t.position)) || seen.has(rel.uri + '|' + t.position)) { skipped++; return; }
    const trackArtist = (t.artists && t.artists.length) ? artistNames_(t.artists) : relArtist;
    const info = lookupBpmKey_(trackArtist, t.title, bpmKey);

    const row = new Array(width).fill('');
    row[cols.artist] = relArtist;
    row[cols.release] = rel.title;
    row[cols.pos] = t.position;
    row[cols.trackArtist] = trackArtist;
    row[cols.title] = t.title;
    row[cols.genre] = g.genre;
    row[cols.tags] = g.tags;
    row[cols.bpm] = info.bpm;
    row[cols.key] = info.key;
    row[cols.camelot] = info.camelot;
    row[cols.label] = cleanName_(lab.name || '');
    row[cols.catno] = lab.catno || '';
    row[cols.year] = rel.year || '';
    row[cols.url] = rel.uri;
    row[cols.source] = info.source;
    rows.push(row);
  });

  if (!rows.length) { ui.alert('All tracks from that release are already in the sheet.'); return; }

  const start = sheet.getLastRow() + 1;
  sheet.getRange(start, 1, rows.length, width).setValues(rows);
  const found = rows.filter(r => r[cols.bpm] !== '').length;
  SpreadsheetApp.getActive().toast(
    `Added ${rows.length} tracks from "${rel.title}" (rows ${start}–${start + rows.length - 1}). ` +
    `BPM/key found for ${found}.` + (skipped ? ` Skipped ${skipped} already in the sheet.` : ''),
    'Crate', 10);
}

// ---------- Fill blanks on existing rows (never overwrites anything you typed) ----------

function fillMissing() {
  const ui = SpreadsheetApp.getUi();
  const bpmKey = prop_('GETSONGBPM_KEY');
  if (!bpmKey) { ui.alert('Add your GetSongBPM key first: Crate → Set API keys…'); return; }

  const sheet = findSheet_();
  const cols = columns_(sheet);
  const n = sheet.getLastRow() - 1;
  if (n < 1) return;
  const values = sheet.getRange(2, 1, n, sheet.getLastColumn()).getValues();

  let tried = 0, filled = 0;
  const changes = [];   // [rowIndex, colIndex, value]
  for (let i = 0; i < values.length && tried < RETRY_LIMIT; i++) {
    const row = values[i];
    const needBpm = row[cols.bpm] === '';
    const needKey = row[cols.key] === '';
    const source = String(row[cols.source]);
    if ((!needBpm && !needKey) || source === 'not found' || source === 'GetSongBPM') continue;
    const artist = row[cols.trackArtist] || row[cols.artist];
    if (!artist || !row[cols.title]) continue;

    tried++;
    const info = lookupBpmKey_(artist, row[cols.title], bpmKey);
    if (info.bpm === '' && info.key === '') { changes.push([i, cols.source, info.source]); continue; }
    filled++;
    if (needBpm && info.bpm !== '') changes.push([i, cols.bpm, info.bpm]);
    if (needKey && info.key) {
      changes.push([i, cols.key, info.key]);
      if (row[cols.camelot] === '') changes.push([i, cols.camelot, info.camelot]);
    }
    changes.push([i, cols.source, 'GetSongBPM']);
  }

  changes.forEach(([i, c, v]) => sheet.getRange(i + 2, c + 1).setValue(v));
  ui.alert(`Looked up ${tried} tracks, filled ${filled}.` +
    (tried >= RETRY_LIMIT ? '\n\nThere are more to do. Run it again to continue.' : '') +
    '\n\nTracks GetSongBPM didn\'t have are marked "not found" in the BPM/Key source column and skipped next time.');
}

// ---------- Genre → first genre + #tags ----------

// Rows that already have Tags are skipped, so this is safe to run more than once
function splitGenres() {
  const sheet = findSheet_();
  const cols = columns_(sheet);
  const n = sheet.getLastRow() - 1;
  if (n < 1) return;
  const gRange = sheet.getRange(2, cols.genre + 1, n, 1);
  const tRange = sheet.getRange(2, cols.tags + 1, n, 1);
  const genres = gRange.getValues(), tags = tRange.getValues();
  let done = 0;
  genres.forEach((r, i) => {
    if (tags[i][0] !== '' || String(r[0]).trim() === '') return;
    const g = genreAndTags_(r[0]);
    genres[i][0] = g.genre;
    tags[i][0] = g.tags;
    done++;
  });
  gRange.setValues(genres);
  tRange.setValues(tags);
  SpreadsheetApp.getUi().alert(`Split ${done} rows. Genre now holds the first genre; Tags holds all of them.`);
}

// "Techno, Dub Techno / Deep" → { genre: 'Techno', tags: '#techno #dubtechno #deep' }
function genreAndTags_(text) {
  const parts = String(text || '').split(/[,\/]/).map(p => p.trim()).filter(Boolean);
  const seen = new Set();
  const tags = [];
  parts.forEach(p => {
    const t = p.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]/g, '');
    if (t && !seen.has(t)) { seen.add(t); tags.push('#' + t); }
  });
  return { genre: parts[0] || '', tags: tags.join(' ') };
}

// ---------- Tidy key notation and fill Camelot ----------

function tidyKeys() {
  const sheet = findSheet_();
  const cols = columns_(sheet);
  const n = sheet.getLastRow() - 1;
  if (n < 1) return;

  const keyRange = sheet.getRange(2, cols.key + 1, n, 1);
  const camRange = sheet.getRange(2, cols.camelot + 1, n, 1);
  const keys = keyRange.getValues();
  const cams = camRange.getValues();
  const keyBg = keyRange.getBackgrounds(), camBg = camRange.getBackgrounds();
  const keyNotes = keyRange.getNotes(), camNotes = camRange.getNotes();

  const result = tidyRows_(keys.map(r => r[0]), cams.map(r => r[0]));
  result.rows.forEach((r, i) => {
    keys[i][0] = r.key;
    cams[i][0] = r.camelot;
    if (r.flagKey) { keyBg[i][0] = FLAG_COLOUR; keyNotes[i][0] = r.note; }
    if (r.flagCamelot) { camBg[i][0] = FLAG_COLOUR; camNotes[i][0] = r.note; }
  });

  keyRange.setValues(keys).setBackgrounds(keyBg).setNotes(keyNotes);
  camRange.setValues(cams).setBackgrounds(camBg).setNotes(camNotes);

  const s = result.stats;
  SpreadsheetApp.getUi().alert(
    `Standardised ${s.renamed} keys (e.g. "B Minor" → "Bm").\n` +
    `Filled Camelot for ${s.camelotFilled} rows.\n` +
    `Filled Key from Camelot for ${s.keyFilled} rows.\n` +
    `Flagged ${s.flagged} cells in orange where Key and Camelot disagree or the key isn't readable. ` +
    `Hover them for details. Nothing flagged was changed.`);
}

// Pure logic (no Sheets calls) so it can be tested outside Apps Script
function tidyRows_(keys, cams) {
  const stats = { renamed: 0, camelotFilled: 0, keyFilled: 0, flagged: 0 };
  const rows = keys.map((rawKey, i) => {
    const rawCam = cams[i];
    const k = parseKey_(rawKey);
    const c = parseCamelot_(rawCam);
    const out = { key: rawKey, camelot: rawCam, flagKey: false, flagCamelot: false, note: '' };
    const keyBlank = rawKey === '' || rawKey === null;
    const camBlank = rawCam === '' || rawCam === null;

    if (k) {
      if (String(rawKey) !== k.key) { out.key = k.key; stats.renamed++; }
      if (camBlank) { out.camelot = k.camelot; stats.camelotFilled++; }
      else if (!c || c.camelot !== k.camelot) {
        out.flagKey = out.flagCamelot = true; stats.flagged += 2;
        out.note = `Crate: Key ${k.key} is ${k.camelot} in Camelot, but the Camelot column says ${rawCam}. Check which is right.`;
      } else if (String(rawCam) !== c.camelot) { out.camelot = c.camelot; }
    } else if (c) {
      if (keyBlank) { out.key = c.key; stats.keyFilled++; }
      else {
        out.flagKey = true; stats.flagged++;
        out.note = `Crate: couldn't read "${rawKey}" as a key. Camelot ${c.camelot} would be ${c.key}.`;
      }
    } else if (!keyBlank) {
      out.flagKey = true; stats.flagged++;
      out.note = `Crate: couldn't read "${rawKey}" as a key.`;
    }
    return out;
  });
  return { rows, stats };
}

// ---------- Key helpers ----------
// Output style matches the sheet: C#, Ab, Bb, F# / Bm, Bbm, F#m, C#m, Abm

const PITCH = { C:0, 'C#':1, Db:1, D:2, 'D#':3, Eb:3, E:4, F:5, 'F#':6, Gb:6, G:7, 'G#':8, Ab:8, A:9, 'A#':10, Bb:10, B:11 };
const NAME = ['C', 'C#', 'D', 'Eb', 'E', 'F', 'F#', 'G', 'Ab', 'A', 'Bb', 'B'];

// "B♭ Minor", "Bbm", "A#m", "C# Major", "F#`" → { key: 'Bbm', camelot: '3A' }
function parseKey_(raw) {
  if (raw === null || raw === undefined || typeof raw === 'number') return null;
  const s = String(raw).replace(/♯/g, '#').replace(/♭/g, 'b').replace(/[`'’.]/g, '').trim();
  const m = /^([A-Ga-g])\s*([#b]?)\s*(major|maj|minor|min|m)?$/i.exec(s);
  if (!m) return null;
  const pc = PITCH[m[1].toUpperCase() + m[2]];
  if (pc === undefined) return null;
  const minor = !!m[3] && /^m(in(or)?)?$/i.test(m[3]);
  return { key: NAME[pc] + (minor ? 'm' : ''), camelot: camelotOf_(pc, minor) };
}

function camelotOf_(pc, minor) {
  const rel = minor ? (pc + 3) % 12 : pc;          // relative major shares the Camelot number
  return ((rel * 7 + 7) % 12 + 1) + (minor ? 'A' : 'B');
}

// "3A" → { camelot: '3A', key: 'Bbm' }
function parseCamelot_(raw) {
  const m = /^\s*(1[0-2]|[1-9])\s*([AB])\s*$/i.exec(String(raw === null ? '' : raw));
  if (!m) return null;
  const camelot = m[1] + m[2].toUpperCase();
  for (let pc = 0; pc < 12; pc++) {
    for (const minor of [false, true]) {
      if (camelotOf_(pc, minor) === camelot) return { camelot, key: NAME[pc] + (minor ? 'm' : '') };
    }
  }
  return null;
}

// ---------- Discogs ----------

function discogs_(path, token, params, attempt) {
  attempt = attempt || 0;
  const q = Object.assign({ token: token }, params || {});
  const qs = Object.keys(q).map(k => encodeURIComponent(k) + '=' + encodeURIComponent(q[k])).join('&');
  const resp = UrlFetchApp.fetch('https://api.discogs.com' + path + '?' + qs, {
    headers: { 'User-Agent': 'CrateSheet/1.0' },
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  if (code === 429 && attempt < 3) { Utilities.sleep(5000); return discogs_(path, token, params, attempt + 1); }
  if (code !== 200) throw new Error('Discogs error ' + code + ': ' + resp.getContentText().slice(0, 200));
  return JSON.parse(resp.getContentText());
}

function parseReleaseId_(input) {
  const m = /discogs\.com\/(?:[^\/]+\/)?release\/(\d+)/i.exec(input) || /^\[?r?(\d{3,})\]?$/i.exec(input);
  return m ? m[1] : null;
}

function searchDiscogs_(input, token) {
  const parts = input.split(/\s+[-–—]\s+/);
  const params = { type: 'release', per_page: 6 };
  if (parts.length >= 2) {
    params.artist = parts[0];
    params.release_title = parts.slice(1).join(' - ');
  } else {
    params.q = input;
  }
  let res = discogs_('/database/search', token, Object.assign({ format: 'Vinyl' }, params));
  if (!res.results || !res.results.length) res = discogs_('/database/search', token, params);
  return res.results || [];
}

function pickCandidate_(candidates, ui) {
  if (candidates.length === 1) return candidates[0].id;
  const lines = candidates.map((c, i) => {
    const label = (c.label && c.label[0]) || '';
    const fmt = (c.format || []).join(', ');
    return `${i + 1}. ${c.title} — ${label} ${c.catno || ''} (${c.year || '?'}, ${c.country || '?'}) [${fmt}]`;
  });
  const r = ui.prompt('Which pressing?', lines.join('\n') + '\n\nEnter a number (blank = 1):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  const n = parseInt(r.getResponseText().trim() || '1', 10);
  return (n >= 1 && n <= candidates.length) ? candidates[n - 1].id : null;
}

function artistNames_(artists) {
  return (artists || []).map((a, i, arr) => {
    const name = cleanName_(a.anv || a.name);
    const join = (i < arr.length - 1) ? (a.join ? ` ${a.join.trim()} ` : ' & ') : '';
    return name + join;
  }).join('').replace(/\s+,/g, ',').trim();
}

// Discogs adds "(2)" etc. to tell apart artists with the same name
function cleanName_(s) { return String(s).replace(/\s+\(\d+\)$/, '').trim(); }

// ---------- GetSongBPM ----------

function lookupBpmKey_(artist, title, apiKey) {
  const blank = source => ({ bpm: '', key: '', camelot: '', source: source });
  if (!apiKey) return blank('no GetSongBPM key yet');

  const lookup = 'song:' + title + ' artist:' + artist;
  const url = 'https://api.getsong.co/search/?api_key=' + encodeURIComponent(apiKey) +
              '&type=both&lookup=' + encodeURIComponent(lookup);
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) return blank('GetSongBPM error ' + resp.getResponseCode());

  const data = JSON.parse(resp.getContentText());
  const results = Array.isArray(data.search) ? data.search : [];
  const want = norm_(artist);
  const match = results.find(s => {
    const got = norm_(s.artist && s.artist.name);
    return want && got && (got.includes(want) || want.includes(got));
  });
  if (!match) return blank('not found');

  let song = match;
  if (!song.tempo && song.id) {
    const r2 = UrlFetchApp.fetch('https://api.getsong.co/song/?api_key=' + encodeURIComponent(apiKey) +
                                 '&id=' + encodeURIComponent(song.id), { muteHttpExceptions: true });
    if (r2.getResponseCode() === 200) song = JSON.parse(r2.getContentText()).song || song;
  }

  const k = parseKey_(song.key_of);
  return {
    bpm: song.tempo ? Number(song.tempo) : '',
    key: k ? k.key : '',
    camelot: k ? k.camelot : '',
    source: 'GetSongBPM'
  };
}

function norm_(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/^the\s+/, '').replace(/[^a-z0-9]/g, '');
}

// ---------- Sheet helpers ----------

// The tab whose header row has "Track Name" and "BPM" (so renaming the tab doesn't break anything)
function findSheet_() {
  const ss = SpreadsheetApp.getActive();
  const hit = ss.getSheets().find(sh => {
    if (sh.getLastColumn() < 1) return false;
    const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v).trim());
    return h.includes(F.title) && h.includes(F.bpm);
  });
  if (!hit) throw new Error('Couldn\'t find a tab with "Track Name" and "BPM" headers in row 1.');
  return hit;
}

// Column positions (0-based) by header name; adds the reference columns at the right if missing
function columns_(sheet) {
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(v => String(v).trim());
  // Tags sits right after Genre
  if (!headers.includes(F.tags) && headers.includes(F.genre)) {
    const gCol = headers.indexOf(F.genre) + 1;
    sheet.insertColumnAfter(gCol);
    sheet.getRange(1, gCol + 1).setValue(F.tags).setFontWeight('bold');
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(v => String(v).trim());
  }
  const missing = Object.values(EXTRA).filter(h => !headers.includes(h));
  if (missing.length) {
    const start = sheet.getLastColumn() + 1;
    sheet.getRange(1, start, 1, missing.length).setValues([missing])
      .setFontWeight('bold').setFontColor('#888888');
    headers = headers.concat(missing);
  }
  const cols = {};
  const all = Object.assign({}, F, EXTRA);
  Object.keys(all).forEach(k => {
    const idx = headers.indexOf(all[k]);
    if (idx < 0) throw new Error('Missing column "' + all[k] + '" in row 1.');
    cols[k] = idx;
  });
  return cols;
}

function existingTracks_(sheet, cols) {
  const set = new Set();
  const n = sheet.getLastRow() - 1;
  if (n < 1) return set;
  sheet.getRange(2, 1, n, sheet.getLastColumn()).getValues().forEach(r => {
    if (r[cols.release]) set.add(trackId_(r[cols.release], r[cols.pos]));
    if (r[cols.url]) set.add(r[cols.url] + '|' + r[cols.pos]);
  });
  return set;
}

function trackId_(release, pos) { return norm_(release) + '|' + String(pos).trim().toUpperCase(); }

function prop_(name) { return PropertiesService.getUserProperties().getProperty(name); }

// Allow local testing with Node; ignored by Apps Script
if (typeof module !== 'undefined') module.exports = { genreAndTags_, parseKey_, parseCamelot_, tidyRows_, camelotOf_ };
