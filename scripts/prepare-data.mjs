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
// 3. Excel (age distribution.xlsx) → age time series per commune
// =====================================================================

const AGE_BANDS = [
  { start: 2,   key: 'age_pct_65plus',    label: 'Aandeel 65+',               labelEn: 'Share 65+',       unit: '%',        div100: true,  yearRow: 2 },
  { start: 28,  key: 'age_pct_under3',    label: 'Aandeel jonger dan 3 jaar',  labelEn: 'Share under 3',   unit: '%',        div100: true,  yearRow: 2 },
  { start: 54,  key: 'age_pct_0_17',      label: 'Aandeel 0-17 jarigen',       labelEn: 'Share 0-17',      unit: '%',        div100: true,  yearRow: 2 },
  { start: 80,  key: 'age_count_12_17',   label: 'Aantal 12-17 jarigen',       labelEn: 'Number 12-17',    unit: 'personen', div100: false, yearRow: 2 },
  { start: 106, key: 'age_pct_18_24',     label: 'Aandeel 18-24 jarigen',      labelEn: 'Share 18-24',     unit: '%',        div100: true,  yearRow: 2 },
  // col 132 skipped: source data contains non-age values despite label
  { start: 158, key: 'age_pct_18_64',     label: 'Aandeel 18-64 jarigen',      labelEn: 'Share 18-64',     unit: '%',        div100: true,  yearRow: 2 },
  { start: 184, key: 'age_pct_30_44',     label: 'Aandeel 30-44 jarigen',      labelEn: 'Share 30-44',     unit: '%',        div100: true,  yearRow: 2 },
  { start: 210, key: 'age_pct_3_5',       label: 'Aandeel 3-5 jarigen',        labelEn: 'Share 3-5',       unit: '%',        div100: true,  yearRow: 2 },
  { start: 236, key: 'age_pct_45_64',     label: 'Aandeel 45-64 jarigen',      labelEn: 'Share 45-64',     unit: '%',        div100: true,  yearRow: 2 },
  { start: 262, key: 'age_pct_6_11',      label: 'Aandeel 6-11 jarigen',       labelEn: 'Share 6-11',      unit: '%',        div100: true,  yearRow: 2 },
  { start: 288, key: 'age_pct_65_79',     label: 'Aandeel 65-79 jarigen',      labelEn: 'Share 65-79',     unit: '%',        div100: true,  yearRow: 2 },
  { start: 314, key: 'age_pct_80plus',    label: 'Aandeel 80 jaar en ouder',   labelEn: 'Share 80+',       unit: '%',        div100: true,  yearRow: 2 },
  { start: 340, key: 'age_count_0_17',    label: 'Aantal 0-17 jarigen',        labelEn: 'Number 0-17',     unit: 'personen', div100: false, yearRow: 1 },
  { start: 366, key: 'age_count_18_64',   label: 'Aantal 18-64 jarigen',       labelEn: 'Number 18-64',    unit: 'personen', div100: false, yearRow: 1 },
  { start: 392, key: 'age_count_65plus',  label: 'Aantal 65 jaar en ouder',    labelEn: 'Number 65+',      unit: 'personen', div100: false, yearRow: 1 },
];

function convertAgeExcel() {
  console.log("\nParsing age distribution.xlsx...");
  const xlsxFile = join(ROOT, "age distribution.xlsx");
  const wb = XLSX.readFile(xlsxFile);
  const ws = wb.Sheets[wb.SheetNames[0]];

  function cell(r, c) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cl = ws[addr];
    return cl ? cl.v : null;
  }

  // Build colMap from band definitions (26 years: 2000–2025)
  const colMap = [];
  for (const band of AGE_BANDS) {
    for (let i = 0; i < 26; i++) {
      const col = band.start + i;
      const year = cell(band.yearRow, col);
      if (typeof year !== 'number') continue;
      colMap.push({ col, key: band.key, year, div100: band.div100 });
    }
  }
  console.log(`  ${colMap.length} (col, key, year) entries`);

  const communes = {};

  // Commune rows start at 7; skip header rows (rows 0-6)
  for (let r = 7; r <= 25; r++) {
    const rawCode = cell(r, 0);
    if (rawCode === null || rawCode === undefined) continue;
    const niscode = String(rawCode).trim().padStart(5, '0');
    if (!/^\d{5}$/.test(niscode)) continue;
    if (!niscode.startsWith('21')) continue;

    const timeseries = {};
    const latest = {};
    for (const band of AGE_BANDS) timeseries[band.key] = {};

    for (const { col, key, year, div100 } of colMap) {
      const raw = cell(r, col);
      if (raw === null || raw === undefined) continue;
      const val = typeof raw === 'number' ? raw : parseFloat(String(raw).replace(',', '.'));
      if (isNaN(val)) continue;
      timeseries[key][year] = div100 ? Math.round(val / 100 * 1e6) / 1e6 : val;
    }

    for (const band of AGE_BANDS) {
      const sortedYears = Object.keys(timeseries[band.key]).map(Number).sort((a, b) => b - a);
      if (sortedYears.length > 0) {
        latest[band.key] = timeseries[band.key][sortedYears[0]];
      }
    }

    communes[niscode] = { timeseries, latest };
  }

  console.log(`  Read ${Object.keys(communes).length} communes`);

  // Metadata: min/max/step per indicator
  const metadata = {};
  const allYearsSet = new Set();

  for (const band of AGE_BANDS) {
    const allValues = [];
    const yearsWithData = new Set();

    for (const c of Object.values(communes)) {
      const ts = c.timeseries[band.key] || {};
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
    else step = 0.001;

    const sortedYears = [...yearsWithData].sort((a, b) => a - b);

    metadata[band.key] = {
      label: band.label,
      labelEn: band.labelEn,
      unit: band.unit,
      years: sortedYears,
      latestYear: sortedYears[sortedYears.length - 1],
      min: Math.floor(min * 1000) / 1000,
      max: Math.ceil(max * 1000) / 1000,
      step,
      count: Object.keys(communes).length
    };
  }

  const allYears = [...allYearsSet].sort((a, b) => a - b);
  console.log(`  Age year range: ${allYears[0]}–${allYears[allYears.length - 1]}`);
  console.log(`  Age indicators: ${Object.keys(metadata).length}`);

  return { communes, metadata, allYears };
}

// =====================================================================
// 4. Merge age data into migration dataset
// =====================================================================
function mergeAgeData(migData, ageData) {
  console.log("\nMerging age data...");
  let merged = 0;

  for (const [niscode, mig] of Object.entries(migData.communes)) {
    const age = ageData.communes[niscode];
    if (!age) continue;
    for (const [key, ts] of Object.entries(age.timeseries)) {
      mig.timeseries[key] = ts;
    }
    Object.assign(mig.latest, age.latest);
    merged++;
  }

  Object.assign(migData.metadata, ageData.metadata);

  const yearSet = new Set(migData.allYears);
  for (const yr of ageData.allYears) yearSet.add(yr);
  migData.allYears = [...yearSet].sort((a, b) => a - b);

  console.log(`  Merged age data for ${merged} communes`);
}

// =====================================================================
// RUN
// =====================================================================
const geojson = await convertSHP();
const data = convertExcel();
const ageData = convertAgeExcel();
mergeAgeData(data, ageData);
writeFileSync(join(ROOT, "data", "brussels-data.json"), JSON.stringify(data));
console.log(`  Updated brussels-data.json: ${Object.keys(data.communes).length} communes, ${Object.keys(data.metadata).length} indicators`);
mergeData(geojson, data);
console.log("\nDone!");
