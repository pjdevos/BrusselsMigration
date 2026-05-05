/**
 * geocode-permits.mjs
 * Geocodes all permit addresses using Nominatim (OSM) and stores lat/lng
 * directly in data/brussels-permits.json.
 *
 * Run once:  node scripts/geocode-permits.mjs
 * Re-run:    skips permits that already have coordinates.
 *
 * Nominatim policy: max 1 req/sec, descriptive User-Agent required.
 */

import fs   from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DATA_FILE = path.join(__dirname, '../data/brussels-permits.json');
const DELAY_MS  = 1150;   // slightly over 1 s to respect Nominatim limit
const USER_AGENT = 'BrusselsMigrationPolicyTool/1.0 contact:naths@tcd.ie';

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

async function geocode(address, commune) {
  const q = encodeURIComponent(`${address}, ${commune}, Brussels, Belgium`);
  const url = `https://nominatim.openstreetmap.org/search?q=${q}&format=json&limit=1&countrycodes=be`;
  const resp = await fetch(url, { headers: { 'User-Agent': USER_AGENT } });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  const data = await resp.json();
  if (data && data.length > 0) {
    return { lat: parseFloat(data[0].lat), lng: parseFloat(data[0].lon) };
  }
  return null;
}

async function main() {
  const raw   = fs.readFileSync(DATA_FILE, 'utf8');
  const data  = JSON.parse(raw);
  const communes = Object.keys(data.byCommune);

  let total = 0, alreadyDone = 0, succeeded = 0, failed = 0;

  for (const commune of communes) {
    const permits = data.byCommune[commune];
    for (const p of permits) {
      total++;
      if (p.lat != null && p.lng != null) { alreadyDone++; continue; }

      process.stdout.write(`[${succeeded + failed + 1}] ${commune} — ${p.address} ... `);
      try {
        const coords = await geocode(p.address, commune);
        if (coords) {
          p.lat = coords.lat;
          p.lng = coords.lng;
          process.stdout.write(`OK (${coords.lat.toFixed(5)}, ${coords.lng.toFixed(5)})\n`);
          succeeded++;
        } else {
          p.lat = null;
          p.lng = null;
          process.stdout.write('NO RESULT\n');
          failed++;
        }
      } catch (e) {
        process.stdout.write(`ERROR: ${e.message}\n`);
        failed++;
      }

      // Save incrementally so progress isn't lost on crash
      fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2));
      await sleep(DELAY_MS);
    }
  }

  console.log(`\nDone. Total: ${total} | Already had coords: ${alreadyDone} | Geocoded: ${succeeded} | Failed: ${failed}`);
}

main().catch(e => { console.error(e); process.exit(1); });
