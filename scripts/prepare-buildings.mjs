/**
 * prepare-buildings.mjs
 * Pre-fetches OSM building footprints for the Brussels Capital Region from
 * Overpass, runs point-in-polygon commune tagging, and writes
 * data/brussels-buildings.json.
 *
 * Run once (or whenever OSM data should be refreshed):
 *   node scripts/prepare-buildings.mjs
 *
 * The browser then loads this static file instantly instead of hitting Overpass.
 */

import { readFileSync, writeFileSync } from 'fs';
import { dirname, join }               from 'path';
import { fileURLToPath }               from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const dataDir   = join(__dirname, '..', 'data');

// ── Load commune polygons for PIP tagging ────────────────────────────────────
const geojson  = JSON.parse(readFileSync(join(dataDir, 'brussels-merged.geojson'), 'utf8'));
const communes = geojson.features.map(f => {
  const geom  = f.geometry;
  const rings = geom.type === 'Polygon'
    ? [geom.coordinates[0]]
    : geom.type === 'MultiPolygon' ? geom.coordinates.map(p => p[0]) : [];
  let minLat = 90, maxLat = -90, minLon = 180, maxLon = -180;
  for (const r of rings) for (const [ln, lt] of r) {
    if (lt < minLat) minLat = lt;  if (lt > maxLat) maxLat = lt;
    if (ln < minLon) minLon = ln;  if (ln > maxLon) maxLon = ln;
  }
  return { niscode: String(f.properties.niscode), rings, minLat, maxLat, minLon, maxLon };
});

function pip(px, py, ring) {
  let inside = false;
  for (let i = 0, j = ring.length - 1; i < ring.length; j = i++) {
    const xi = ring[i][0], yi = ring[i][1], xj = ring[j][0], yj = ring[j][1];
    if (((yi > py) !== (yj > py)) && px < (xj - xi) * (py - yi) / (yj - yi) + xi)
      inside = !inside;
  }
  return inside;
}

// ── Fetch from Overpass ──────────────────────────────────────────────────────
// Full BCR extent covers all 19 communes including Uccle/Watermael/Woluwe.
const q = `[out:json][timeout:120];(way["building"](50.755,4.235,50.925,4.495););out body geom qt;`;

console.log('Fetching OSM buildings from Overpass (this takes ~30–60 s)…');
const t0       = Date.now();
const response = await fetch('https://overpass-api.de/api/interpreter', {
  method:  'POST',
  headers: {
    'Content-Type': 'application/x-www-form-urlencoded',
    'User-Agent':   'BrusselsMigration/1.0 (build script; contact naths@tcd.ie)',
    'Accept':       'application/json',
  },
  body: 'data=' + encodeURIComponent(q),
});
if (!response.ok) throw new Error(`Overpass HTTP ${response.status}`);
const osm = await response.json();
console.log(`Overpass returned ${osm.elements.length} elements in ${((Date.now()-t0)/1000).toFixed(1)} s`);

// ── Convert OSM ways → GeoJSON features ─────────────────────────────────────
const features = [];
for (const el of (osm.elements || [])) {
  if (el.type !== 'way' || !el.geometry || el.geometry.length < 4) continue;
  const coords = el.geometry.map(n => [n.lon, n.lat]);
  // Close ring if not already closed
  const fst = coords[0], lst = coords[coords.length - 1];
  if (fst[0] !== lst[0] || fst[1] !== lst[1]) coords.push([...fst]);

  const tags    = el.tags || {};
  let   h       = 8, hTag = false;
  if (tags.height) {
    h    = parseFloat(String(tags.height).replace(/[^\d.]/g, '')) || 8;
    hTag = true;
  } else if (tags['building:levels']) {
    h = (parseInt(tags['building:levels']) || 2) * 3.2;
  }

  const houseNum = tags['addr:housenumber'] || '';
  const streetNm = tags['addr:street']      || '';
  const street   = (streetNm + (houseNum ? ' ' + houseNum : '')).trim() || null;

  features.push({
    type: 'Feature',
    geometry: { type: 'Polygon', coordinates: [coords] },
    properties: {
      h, minH: 0, hTag,
      levels:  tags['building:levels'] ? (parseInt(tags['building:levels']) || null) : null,
      bType:   tags.building   || null,
      street,
      bName:   tags.name || tags['name:fr'] || tags['name:nl'] || tags['name:en']
             || tags.brand || tags.operator || null,
      amenity: tags.amenity  || null,
      shop:    tags.shop     || null,
      office:  tags.office   || null,
      tourism: tags.tourism  || null,
      leisure: tags.leisure  || null,
      osmId:   el.id         || null,
    },
  });
}

// ── Point-in-polygon commune tagging ────────────────────────────────────────
let tagged = 0;
for (const feat of features) {
  const ring = feat.geometry.coordinates[0];
  let cx = 0, cy = 0;
  for (const [ln, lt] of ring) { cx += ln; cy += lt; }
  cx /= ring.length; cy /= ring.length;

  for (const c of communes) {
    if (cy < c.minLat || cy > c.maxLat || cx < c.minLon || cx > c.maxLon) continue;
    for (const r of c.rings) {
      if (pip(cx, cy, r)) { feat.properties.niscode = c.niscode; break; }
    }
    if (feat.properties.niscode) break;
  }
  if (feat.properties.niscode) tagged++;
}
console.log(`Tagged ${tagged} / ${features.length} buildings with commune niscode`);

// ── Write one file per commune ────────────────────────────────────────────────
// Browser lazy-loads only the selected commune (~2 MB gzipped each)
// instead of a 150 MB monolithic file.
import { mkdirSync } from 'fs';

const bldDir = join(dataDir, 'buildings');
mkdirSync(bldDir, { recursive: true });

// Group by niscode
const byNis = {};
for (const feat of features) {
  const nis = feat.properties.niscode;
  if (!nis) continue;
  (byNis[nis] = byNis[nis] || []).push(feat);
}

// Reduce coordinate precision to 5 dp (≈ 1 m accuracy — more than enough for buildings)
function roundCoords(geoJson) {
  for (const feat of geoJson.features) {
    feat.geometry.coordinates[0] = feat.geometry.coordinates[0]
      .map(([ln, lt]) => [Math.round(ln * 1e5) / 1e5, Math.round(lt * 1e5) / 1e5]);
  }
}

let totalWritten = 0;
for (const [nis, feats] of Object.entries(byNis)) {
  const geoJson = { type: 'FeatureCollection', features: feats };
  roundCoords(geoJson);
  const path = join(bldDir, `buildings-${nis}.json`);
  writeFileSync(path, JSON.stringify(geoJson));
  console.log(`  ${nis}: ${feats.length} buildings → ${path}`);
  totalWritten += feats.length;
}

// Write a manifest so the browser knows which commune files exist
const manifest = Object.fromEntries(Object.entries(byNis).map(([nis, f]) => [nis, f.length]));
writeFileSync(join(bldDir, 'manifest.json'), JSON.stringify(manifest, null, 2));
console.log(`✓ Written ${totalWritten} buildings across ${Object.keys(byNis).length} communes → ${bldDir}`);
