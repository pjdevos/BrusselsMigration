import { readFileSync, writeFileSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';
import XLSX from 'xlsx';
import proj4 from 'proj4';
import * as shapefile from 'shapefile';

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = join(__dirname, '..');

// =====================================================================
// EPSG:31370 (Belgian Lambert 72) → EPSG:4326 projection
// +towgs84 is mandatory; without it polygons land ~100 m off WGS84.
// =====================================================================
proj4.defs("EPSG:31370",
  "+proj=lcc +lat_1=51.16666723333333 +lat_2=49.8333339 +lat_0=90 " +
  "+lon_0=4.367486666666666 +x_0=150000.013 +y_0=5400088.438 " +
  "+ellps=intl +towgs84=-106.869,52.2978,-103.724,0.3366,-0.457,1.8422,-1.2747 " +
  "+units=m +no_defs"
);

function reproject(x, y) {
  const [lng, lat] = proj4("EPSG:31370", "EPSG:4326", [x, y]);
  return [Math.round(lng * 1e6) / 1e6, Math.round(lat * 1e6) / 1e6];
}

function reprojectRing(ring) {
  return ring.map(([x, y]) => reproject(x, y));
}

function reprojectGeometry(geom) {
  if (geom.type === "Polygon") {
    return { type: "Polygon", coordinates: geom.coordinates.map(reprojectRing) };
  }
  if (geom.type === "MultiPolygon") {
    return {
      type: "MultiPolygon",
      coordinates: geom.coordinates.map(poly => poly.map(reprojectRing))
    };
  }
  return geom;
}

// =====================================================================
// 1. Shapefile → GeoJSON (19 Brussels communes)
// =====================================================================
async function convertSHP() {
  console.log("Parsing Shapefile...");
  const shpPath = join(ROOT, "UrbISAdminUnits_31370_SHP_04000_20251017", "shp", "UrbISAdminUnits_04000_Municipalities.shp");
  const dbfPath = join(ROOT, "UrbISAdminUnits_31370_SHP_04000_20251017", "shp", "UrbISAdminUnits_04000_Municipalities.dbf");

  const fc = await shapefile.read(shpPath, dbfPath);
  console.log(`  Found ${fc.features.length} features`);

  const features = [];
  for (const f of fc.features) {
    const props = f.properties || {};
    const niscode = String(props.NISCODE).trim().padStart(5, "0");
    if (niscode === "04000") continue; // Region total — defensive
    if (!niscode.startsWith("21")) continue; // Brussels communes only

    const nameFr = String(props.NAMEFRE || "").trim();
    const nameDut = String(props.NAMEDUT || "").trim();
    const area = props.AREA != null ? Number(props.AREA) : null;

    const geometry = reprojectGeometry(f.geometry);

    features.push({
      type: "Feature",
      properties: {
        niscode,
        nameFr,
        nameDut,
        name: nameDut || nameFr,
        area
      },
      geometry
    });
  }

  const geojson = { type: "FeatureCollection", features };
  writeFileSync(join(ROOT, "data", "brussels-communes.geojson"), JSON.stringify(geojson));
  console.log(`  Wrote ${features.length} commune features`);

  const test = features.find(f => f.properties.niscode === "21004");
  if (test) {
    const firstCoord = test.geometry.type === "Polygon"
      ? test.geometry.coordinates[0][0]
      : test.geometry.coordinates[0][0][0];
    console.log(`  Verification — Brussels (21004): first coord [${firstCoord}]`);
  }

  return geojson;
}

// =====================================================================
// 2. Excel (Brussels Migration.xlsx) → JSON with full time series
// =====================================================================

const INDICATORS = [
  { key: "total_population",         label: "Totale bevolking",            labelEn: "Total population",          unit: "inwoners" },
  { key: "internal_in_num",          label: "Interne migratie IN",         labelEn: "Internal migration IN",     unit: "personen" },
  { key: "internal_in_pct",          label: "Interne migratie IN",         labelEn: "Internal migration IN",     unit: "%" },
  { key: "internal_out_num",         label: "Interne migratie UIT",        labelEn: "Internal migration OUT",    unit: "personen" },
  { key: "internal_out_pct",         label: "Interne migratie UIT",        labelEn: "Internal migration OUT",    unit: "%" },
  { key: "internal_balance_num",     label: "Interne migratie saldo",      labelEn: "Internal migration balance", unit: "personen" },
  { key: "internal_balance_pct",     label: "Interne migratie saldo",      labelEn: "Internal migration balance", unit: "%" },
  { key: "international_balance_num", label: "Internationale migratie saldo", labelEn: "International migration balance", unit: "personen" },
  { key: "international_balance_pct", label: "Internationale migratie saldo", labelEn: "International migration balance", unit: "%" },
];
const INDICATOR_BY_KEY = Object.fromEntries(INDICATORS.map(i => [i.key, i]));

function convertExcel() {
  console.log("\nParsing Brussels Migration.xlsx...");
  const xlsxFile = join(ROOT, "Brussels Migration.xlsx");
  const wb = XLSX.readFile(xlsxFile);
  const sheetName = wb.SheetNames.includes("2000-2024") ? "2000-2024" : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const range = XLSX.utils.decode_range(ws["!ref"]);

  function cell(r, c) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cl = ws[addr];
    return cl ? cl.v : null;
  }

  // Build flat colMap = [{ col, key, year }, ...] from the verified band layout.
  // Years run 2024 down to 2000 (descending) in every band.
  const colMap = [];
  const expectedYears = [];
  for (let y = 2024; y >= 2000; y--) expectedYears.push(y);

  // Band 1: TOTAL POPULATION — cols 3..27 (25 cols, year only)
  for (let i = 0; i < 25; i++) {
    colMap.push({ col: 3 + i, key: "total_population", year: expectedYears[i] });
  }
  // Bands 2-5: paired (num, pct), starting at given offsets
  const pairedBands = [
    { start: 29, numKey: "internal_in_num",          pctKey: "internal_in_pct" },
    { start: 79, numKey: "internal_out_num",         pctKey: "internal_out_pct" },
    { start: 129, numKey: "internal_balance_num",    pctKey: "internal_balance_pct" },
    { start: 179, numKey: "international_balance_num", pctKey: "international_balance_pct" },
  ];
  for (const band of pairedBands) {
    for (let i = 0; i < 25; i++) {
      colMap.push({ col: band.start + i * 2,     key: band.numKey, year: expectedYears[i] });
      colMap.push({ col: band.start + i * 2 + 1, key: band.pctKey, year: expectedYears[i] });
    }
  }

  // Validate years against row 3 — abort loudly on mismatch
  for (const entry of colMap) {
    const yr = cell(3, entry.col);
    if (yr !== entry.year) {
      throw new Error(`Year mismatch at col ${entry.col}: expected ${entry.year}, got ${yr}`);
    }
  }
  console.log(`  ${colMap.length} (col, key, year) entries validated against row 3`);

  // Iterate commune rows (4..23). Skip 04000 (Brussels Region total).
  const communes = {};

  for (let r = 4; r <= range.e.r; r++) {
    const rawCode = cell(r, 0);
    if (rawCode === null || rawCode === undefined || rawCode === "") continue;
    const niscode = String(rawCode).trim().padStart(5, "0");
    if (!/^\d{5}$/.test(niscode)) continue; // skip footer rows ("Bron: ...")
    if (niscode === "04000") continue; // skip region total
    if (!niscode.startsWith("21")) continue; // Brussels communes only

    const nameFr = String(cell(r, 1) || "").trim();
    const nameDut = String(cell(r, 2) || "").trim();

    const timeseries = {};
    const latest = {};
    for (const ind of INDICATORS) timeseries[ind.key] = {};

    for (const { col, key, year } of colMap) {
      const val = cell(r, col);
      if (val === null || val === undefined || val === "") continue;
      if (typeof val !== "number") continue; // skip non-numeric (e.g. "ND", "VS")
      timeseries[key][year] = val;
    }

    for (const ind of INDICATORS) {
      const sortedYears = Object.keys(timeseries[ind.key]).map(Number).sort((a, b) => b - a);
      if (sortedYears.length > 0) {
        latest[ind.key] = timeseries[ind.key][sortedYears[0]];
      }
    }

    communes[niscode] = {
      code: niscode,
      nameFr,
      nameDut,
      name: nameDut || nameFr,
      timeseries,
      latest
    };
  }

  console.log(`  Read ${Object.keys(communes).length} communes`);

  // Metadata: min/max/step per indicator across all years
  const metadata = {};
  const allYearsSet = new Set();

  for (const ind of INDICATORS) {
    const allValues = [];
    const yearsWithData = new Set();

    for (const c of Object.values(communes)) {
      const ts = c.timeseries[ind.key] || {};
      for (const [yr, val] of Object.entries(ts)) {
        allValues.push(val);
        yearsWithData.add(Number(yr));
        allYearsSet.add(Number(yr));
      }
    }

    if (allValues.length === 0) continue;

    const min = Math.min(...allValues);
    const max = Math.max(...allValues);
    const range_ = max - min;
    let step;
    if (range_ > 10000) step = 100;
    else if (range_ > 1000) step = 50;
    else if (range_ > 100) step = 5;
    else if (range_ > 10) step = 0.5;
    else if (range_ > 1) step = 0.1;
    else step = 0.001; // for percent values stored as decimal fractions

    const sortedYears = [...yearsWithData].sort((a, b) => a - b);

    metadata[ind.key] = {
      label: ind.label,
      labelEn: ind.labelEn,
      unit: ind.unit,
      years: sortedYears,
      latestYear: sortedYears[sortedYears.length - 1],
      min: Math.floor(min * 1000) / 1000,
      max: Math.ceil(max * 1000) / 1000,
      step,
      count: Object.keys(communes).length
    };
  }

  const allYears = [...allYearsSet].sort((a, b) => a - b);
  console.log(`  Global year range: ${allYears[0]}–${allYears[allYears.length - 1]}`);
  console.log(`  Indicators: ${Object.keys(metadata).length}`);

  const result = { communes, metadata, allYears };
  const outPath = join(ROOT, "data", "brussels-data.json");
  writeFileSync(outPath, JSON.stringify(result));
  console.log(`  Wrote ${Object.keys(communes).length} communes to ${outPath}`);

  return result;
}

// =====================================================================
// 3. Merge: embed LATEST values into GeoJSON properties (join on niscode)
// =====================================================================
function mergeData(geojson, data) {
  console.log("\nMerging data into GeoJSON...");
  let matched = 0, unmatched = 0;

  for (const feature of geojson.features) {
    const niscode = feature.properties.niscode;
    const c = data.communes[niscode];
    if (c) {
      feature.properties.naam = c.nameDut || c.nameFr;
      Object.assign(feature.properties, c.latest);
      feature.properties.hasData = true;
      matched++;
    } else {
      feature.properties.hasData = false;
      unmatched++;
    }
  }

  const outPath = join(ROOT, "data", "brussels-merged.geojson");
  writeFileSync(outPath, JSON.stringify(geojson));
  console.log(`  Matched: ${matched}, No data: ${unmatched}`);
}

// =====================================================================
// RUN
// =====================================================================
const geojson = await convertSHP();
const data = convertExcel();
mergeData(geojson, data);
console.log("\nDone!");
