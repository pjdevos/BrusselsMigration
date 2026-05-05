import { readFileSync, writeFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));

function parseCSVFull(content) {
  content = content.replace(/^\uFEFF/, '');
  const rows = [];
  let i = 0, n = content.length;
  let row = [], field = '', inQ = false;
  while (i < n) {
    const ch = content[i];
    if (ch === '"') {
      if (inQ && content[i + 1] === '"') { field += '"'; i += 2; continue; }
      inQ = !inQ; i++; continue;
    }
    if (ch === ',' && !inQ) { row.push(field.trim()); field = ''; i++; continue; }
    if ((ch === '\n' || ch === '\r') && !inQ) {
      if (ch === '\r' && content[i + 1] === '\n') i++;
      row.push(field.trim()); field = '';
      if (row.some(v => v)) rows.push(row);
      row = []; i++; continue;
    }
    field += ch; i++;
  }
  if (field || row.length) { row.push(field.trim()); if (row.some(v => v)) rows.push(row); }
  return rows;
}

function csvToObjects(rows) {
  const headers = rows[0];
  return rows.slice(1).map(r => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = r[i] || '');
    return obj;
  });
}

const BRUSSELS_COMMUNES = new Set([
  'Anderlecht', 'Auderghem', 'Berchem-Sainte-Agathe', 'Bruxelles',
  'Etterbeek', 'Evere', 'Forest', 'Ganshoren', 'Ixelles', 'Jette',
  'Koekelberg', 'Molenbeek-Saint-Jean', 'Saint-Gilles', 'Saint-Josse-ten-Noode',
  'Schaerbeek', 'Uccle', 'Watermael-Boitsfort', 'Woluwe-Saint-Lambert', 'Woluwe-Saint-Pierre'
]);

const dataDir = join(__dirname, '..', 'data');

// Parse both CSVs
const openPermitsRaw = csvToObjects(parseCSVFull(readFileSync(join(dataDir, 'Register of Open Permits.csv'), 'utf8')));
const mixedLicensesRaw = csvToObjects(parseCSVFull(readFileSync(join(dataDir, 'Mixed Licenses.csv'), 'utf8')));

function normalizePermit(row, source) {
  return {
    ref: row['Référence'] || '',
    address: row['Adresse'] || '',
    commune: row['Commune'] || '',
    submitted: row['Dépôt'] ? row['Dépôt'].split(' ')[0] : '',
    complete: row['Dossier complet'] ? row['Dossier complet'].split(' ')[0] : '',
    decision: row['Décision'] ? row['Décision'].split(' ')[0] : '',
    status: row['Statut'] || '',
    description: row['Objet'] || '',
    authority: row['Autorité délivrante'] || '',
    charges: row['Charges'] ? parseFloat(row['Charges']) || null : null,
    source,
  };
}

const allPermits = [
  ...openPermitsRaw.map(r => normalizePermit(r, 'open')),
  ...mixedLicensesRaw.map(r => normalizePermit(r, 'mixed')),
].filter(p => BRUSSELS_COMMUNES.has(p.commune));

// Group by commune
const byCommune = {};
for (const p of allPermits) {
  if (!byCommune[p.commune]) byCommune[p.commune] = [];
  byCommune[p.commune].push(p);
}

// Sort each commune's permits by submission date descending
for (const c of Object.keys(byCommune)) {
  byCommune[c].sort((a, b) => (b.submitted || '').localeCompare(a.submitted || ''));
}

const output = {
  total: allPermits.length,
  communes: Object.keys(byCommune).length,
  byCommune,
};

const outPath = join(dataDir, 'brussels-permits.json');
writeFileSync(outPath, JSON.stringify(output, null, 2));
console.log(`✓ Written ${allPermits.length} permits across ${Object.keys(byCommune).length} communes → ${outPath}`);
