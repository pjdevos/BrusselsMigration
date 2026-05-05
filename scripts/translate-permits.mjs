/**
 * translate-permits.mjs
 * Translates permit descriptions from French → English using the
 * unofficial Google Translate endpoint (no API key required).
 *
 * Adds a `descriptionEn` field to every permit in data/brussels-permits.json.
 * Skips permits that already have `descriptionEn` set.
 * Saves incrementally after each translation so progress survives crashes.
 *
 * Run once:  node scripts/translate-permits.mjs
 */

import fs   from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname  = path.dirname(fileURLToPath(import.meta.url));
const DATA_FILE  = path.join(__dirname, '../data/brussels-permits.json');
const DELAY_MS   = 600;  // stay well under rate limits

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

async function translateFrEn(text) {
  if (!text || !text.trim()) return '';
  const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=fr&tl=en&dt=t&q=${encodeURIComponent(text)}`;
  const resp = await fetch(url, {
    headers: { 'User-Agent': 'Mozilla/5.0 (compatible; BrusselsMigrationTool/1.0)' }
  });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  const data = await resp.json();
  // Response is nested arrays: [[["translated","original",...], ...], ...]
  return data[0].map(chunk => chunk[0]).join('').trim();
}

async function main() {
  const raw  = fs.readFileSync(DATA_FILE, 'utf8');
  const data = JSON.parse(raw);

  let total = 0, skipped = 0, done = 0, failed = 0;

  for (const [commune, perms] of Object.entries(data.byCommune)) {
    for (const p of perms) {
      total++;
      if (p.descriptionEn !== undefined) { skipped++; continue; }
      if (!p.description || !p.description.trim()) {
        p.descriptionEn = '';
        skipped++;
        continue;
      }

      process.stdout.write(`[${done + failed + 1}] ${commune} — ${p.address.slice(0, 40)} ... `);
      try {
        const en = await translateFrEn(p.description);
        p.descriptionEn = en;
        process.stdout.write(`OK\n`);
        done++;
      } catch (e) {
        process.stdout.write(`ERROR: ${e.message}\n`);
        p.descriptionEn = '';   // mark as attempted so we don't retry forever
        failed++;
      }

      // Save after every translation so progress is never lost
      fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2));
      await sleep(DELAY_MS);
    }
  }

  console.log(`\nDone. Total: ${total} | Skipped (already done): ${skipped} | Translated: ${done} | Failed: ${failed}`);
}

main().catch(e => { console.error(e); process.exit(1); });
