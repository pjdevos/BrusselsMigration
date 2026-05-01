import { readFileSync, writeFileSync, existsSync } from 'fs';
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
  const xlsxFile = existsSync(join(ROOT, 'data', 'Brussels Migration.xlsx')) ? join(ROOT, 'data', 'Brussels Migration.xlsx') : join(ROOT, 'Brussels Migration.xlsx');
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
  const xlsxFile = existsSync(join(ROOT, 'data', 'age distribution.xlsx')) ? join(ROOT, 'data', 'age distribution.xlsx') : join(ROOT, 'age distribution.xlsx');
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
// 4. Generic merge helper — adds any extra dataset into the main one
// =====================================================================
function mergeDataset(main, extra, label) {
  if (!extra) return;
  console.log(`\nMerging ${label}...`);
  let merged = 0;
  for (const [niscode, mainC] of Object.entries(main.communes)) {
    const extraC = extra.communes[niscode];
    if (!extraC) continue;
    for (const [key, ts] of Object.entries(extraC.timeseries)) {
      mainC.timeseries[key] = ts;
    }
    Object.assign(mainC.latest, extraC.latest);
    merged++;
  }
  Object.assign(main.metadata, extra.metadata);
  const yearSet = new Set(main.allYears);
  for (const yr of extra.allYears) yearSet.add(yr);
  main.allYears = [...yearSet].sort((a, b) => a - b);
  console.log(`  Merged ${label} for ${merged} communes`);
}

// =====================================================================
// 5. CSV (SPF Finances BuildingState, annual 2011–2025) → housing indicators
//    Source: finances.belgium.be — Characteristics of cadastral plots — Condition of the building
//    Files: MunicipalityWideBuildingState_YYYYMMDD.csv (one per year)
// =====================================================================

const BS_VARS = [
  { key: 'hq_qual_basic',       label: 'Eenvoudige kwaliteit',     labelEn: 'Basic quality',        unit: '%' },
  { key: 'hq_qual_normal',      label: 'Normale kwaliteit',        labelEn: 'Normal quality',       unit: '%' },
  { key: 'hq_qual_luxurious',   label: 'Luxueuze kwaliteit',       labelEn: 'Luxurious quality',    unit: '%' },
  { key: 'hq_build_pre1900',    label: 'Bouw vóór 1900',           labelEn: 'Built pre-1900',       unit: '%' },
  { key: 'hq_build_1900_1940',  label: 'Bouw 1900–1940',           labelEn: 'Built 1900–1940',      unit: '%' },
  { key: 'hq_build_1941_1960',  label: 'Bouw 1941–1960',           labelEn: 'Built 1941–1960',      unit: '%' },
  { key: 'hq_build_1961_1990',  label: 'Bouw 1961–1990',           labelEn: 'Built 1961–1990',      unit: '%' },
  { key: 'hq_build_1991_2010',  label: 'Bouw 1991–2010',           labelEn: 'Built 1991–2010',      unit: '%' },
  { key: 'hq_build_post2011',   label: 'Bouw 2011 en later',       labelEn: 'Built 2011+',          unit: '%' },
  { key: 'hq_renov_by1982',     label: 'Gerenoveerd t.e.m. 1981',  labelEn: 'Renovated by 1981',    unit: '%' },
  { key: 'hq_renov_by1992',     label: 'Gerenoveerd t.e.m. 1991',  labelEn: 'Renovated by 1991',    unit: '%' },
  { key: 'hq_renov_by2002',     label: 'Gerenoveerd t.e.m. 2001',  labelEn: 'Renovated by 2001',    unit: '%' },
  { key: 'hq_renov_by2012',     label: 'Gerenoveerd t.e.m. 2011',  labelEn: 'Renovated by 2011',    unit: '%' },
  { key: 'hq_renov_by2022',     label: 'Gerenoveerd t.e.m. 2021',  labelEn: 'Renovated by 2021',    unit: '%' },
  { key: 'hq_type_closed',      label: 'Gesloten bebouwing',       labelEn: 'Closed type',          unit: '%' },
  { key: 'hq_type_halfopen',    label: 'Halfopen bebouwing',       labelEn: 'Half-open type',       unit: '%' },
  { key: 'hq_type_open',        label: 'Open bebouwing',           labelEn: 'Open type',            unit: '%' },
];

function convertBuildingState() {
  // Try local copy first, then absolute Downloads path
  const bsDirs = [
    join(ROOT, 'data', 'Characteristics Cadastral Data'),
    join(ROOT, 'building_state'),
    'C:/Users/sushe/Downloads/Datasets/Characteristics Cadastral Data',
    'C:/Users/sushe/Downloads/Characteristics Cadastral Data',
  ];
  let bsDir = null;
  for (const d of bsDirs) { if (existsSync(d)) { bsDir = d; break; } }
  if (!bsDir) {
    console.log('\nSkipping BuildingState CSVs — folder not found, falling back to house_quality.xlsx');
    return convertHouseQualityFallback();
  }
  console.log(`\nParsing BuildingState CSVs from ${bsDir}...`);

  const YEARS = [2011,2012,2013,2014,2015,2016,2017,2018,2019,2020,2021,2022,2023,2024,2025];
  const commData = {};  // { niscode: { timeseries: { varKey: { year: val } } } }

  for (const year of YEARS) {
    const csvPath = join(bsDir, String(year), `MunicipalityWideBuildingState_${year}0101.csv`);
    if (!existsSync(csvPath)) { console.log(`  Skipping ${year} — file not found`); continue; }

    const raw = readFileSync(csvPath, 'utf8').replace(/^\uFEFF/, '');
    const lines = raw.split('\n');
    const hdr = lines[0].split(';');
    const hi = c => hdr.indexOf(c);

    // Group rows by NISCode for Brussels (21xxx)
    const byNis = {};
    for (const line of lines.slice(1)) {
      if (!line.trim()) continue;
      const cols = line.split(';');
      const nis = cols[0];
      if (!nis.startsWith('21')) continue;
      if (!byNis[nis]) byNis[nis] = [];
      byNis[nis].push(cols);
    }

    for (const [nis, rows] of Object.entries(byNis)) {
      const s = col => rows.reduce((acc, c) => acc + (parseInt(c[hi(col)]) || 0), 0);

      // Quality
      const lux = s('LuxuriousQuality'), norm = s('NormalQuality'),
            basic = s('BasicQuality'),   naQ  = s('NAQuality');
      const qualTotal = lux + norm + basic + naQ;

      // Building age
      const b1849=s('Built1849'), b1850=s('Built1850'), b1875=s('Built1875');
      const b1900=s('Built1900'), b1919=s('Built1919'), b1931=s('Built1931');
      const b1941=s('Built1941'), b1951=s('Built1951');
      const b1961=s('Built1961'), b1971=s('Built1971'), b1981=s('Built1981');
      const b1991=s('Built1991'), b2001=s('Built2001');
      const b2011=s('Built2011'), b2021=s('Built2021');
      const builtTotal = b1849+b1850+b1875+b1900+b1919+b1931+b1941+b1951+
                         b1961+b1971+b1981+b1991+b2001+b2011+b2021;

      // Renovation
      const r1982=s('Renewed1982'), r1992=s('Renewed1992'), r2002=s('Renewed2002'),
            r2012=s('Renewed2012'), r2022=s('Renewed2022'), naR=s('NARenewed');
      const renewTotal = r1982+r1992+r2002+r2012+r2022+naR;

      // Construction type
      const tClosed=s('ClosedConstructionType'), tHalf=s('HalfOpenConstructionType'),
            tOpen=s('OpenConstructionType'),      naT=s('NAConstructionType');
      const typeTotal = tClosed+tHalf+tOpen+naT;

      const pct = (num, den) => den > 0 ? Math.round(num / den * 1e6) / 1e6 : null;

      const vals = {
        hq_qual_basic:      pct(basic,   qualTotal),
        hq_qual_normal:     pct(norm,    qualTotal),
        hq_qual_luxurious:  pct(lux,     qualTotal),
        hq_build_pre1900:   pct(b1849+b1850+b1875,           builtTotal),
        hq_build_1900_1940: pct(b1900+b1919+b1931,           builtTotal),
        hq_build_1941_1960: pct(b1941+b1951,                 builtTotal),
        hq_build_1961_1990: pct(b1961+b1971+b1981,           builtTotal),
        hq_build_1991_2010: pct(b1991+b2001,                 builtTotal),
        hq_build_post2011:  pct(b2011+b2021,                 builtTotal),
        hq_renov_by1982:    pct(r1982, renewTotal),
        hq_renov_by1992:    pct(r1992, renewTotal),
        hq_renov_by2002:    pct(r2002, renewTotal),
        hq_renov_by2012:    pct(r2012, renewTotal),
        hq_renov_by2022:    pct(r2022, renewTotal),
        hq_type_closed:     pct(tClosed, typeTotal),
        hq_type_halfopen:   pct(tHalf,   typeTotal),
        hq_type_open:       pct(tOpen,   typeTotal),
      };

      if (!commData[nis]) {
        commData[nis] = { timeseries: {} };
        for (const v of BS_VARS) commData[nis].timeseries[v.key] = {};
      }
      for (const [k, v] of Object.entries(vals)) {
        if (v !== null) commData[nis].timeseries[k][year] = v;
      }
    }
    console.log(`  ${year}: processed ${Object.keys(byNis).length} Brussels communes`);
  }

  // Build latest + communes output
  const out = {};
  for (const [nis, cd] of Object.entries(commData)) {
    const latest = {};
    for (const v of BS_VARS) {
      const ts = cd.timeseries[v.key];
      const yrs = Object.keys(ts).map(Number).sort((a,b) => b-a);
      if (yrs.length) latest[v.key] = ts[yrs[0]];
    }
    out[nis] = { timeseries: cd.timeseries, latest };
  }

  // Build metadata
  const allYears = YEARS.filter(y =>
    Object.values(out).some(c => Object.keys(c.timeseries['hq_qual_basic'] || {}).includes(String(y)))
  );
  const metadata = {};
  for (const v of BS_VARS) {
    const vals = Object.values(out).map(c => c.latest[v.key]).filter(x => x != null);
    if (!vals.length) continue;
    const min = Math.min(...vals), max = Math.max(...vals);
    const range_ = max - min;
    metadata[v.key] = {
      label: v.label, labelEn: v.labelEn, unit: v.unit,
      years: allYears, latestYear: allYears[allYears.length - 1],
      min: Math.floor(min * 1000) / 1000,
      max: Math.ceil(max  * 1000) / 1000,
      step: range_ > 0.1 ? 0.001 : 0.0001,
      count: Object.keys(out).length,
    };
  }

  console.log(`  BuildingState: ${allYears.length} years (${allYears[0]}–${allYears[allYears.length-1]}), ${Object.keys(metadata).length} indicators`);
  return { communes: out, metadata, allYears };
}

// Fallback: original single-year Excel reader (used if CSVs are not found)
function convertHouseQualityFallback() {
  const xlsxFile = existsSync(join(ROOT, 'data', 'house_quality.xlsx')) ? join(ROOT, 'data', 'house_quality.xlsx') : join(ROOT, 'house_quality.xlsx');
  if (!existsSync(xlsxFile)) { console.log('  house_quality.xlsx also not found — skipping'); return null; }
  console.log("  Falling back to house_quality.xlsx (2024 snapshot)...");
  const wb = XLSX.readFile(xlsxFile);
  const ws = wb.Sheets[wb.SheetNames[0]];
  function cell(r, c) { const a = XLSX.utils.encode_cell({r,c}); return ws[a] ? ws[a].v : null; }
  const STATIC_YEAR = 2024;
  const COLS = BS_VARS.map((v, i) => {
    const map = { hq_build_pre1900:2, hq_build_1900_1940:6, hq_build_1941_1960:9,
                  hq_build_1961_1990:13, hq_build_1991_2010:16, hq_build_post2011:19,
                  hq_renov_by1982:22, hq_renov_by1992:23, hq_renov_by2002:24,
                  hq_renov_by2012:25, hq_renov_by2022:26,
                  hq_qual_luxurious:29, hq_qual_normal:30, hq_qual_basic:31,
                  hq_type_closed:34, hq_type_halfopen:35, hq_type_open:36 };
    return { ...v, col: map[v.key] };
  });
  const communes = {};
  for (let r = 2; r <= 20; r++) {
    const rawCode = cell(r, 0);
    if (!rawCode) continue;
    const nis = String(rawCode).trim().padStart(5,'0');
    if (!nis.startsWith('21')) continue;
    const timeseries = {}, latest = {};
    for (const c of COLS) timeseries[c.key] = {};
    for (const { col, key } of COLS) {
      const raw = cell(r, col);
      if (raw == null) continue;
      const val = typeof raw === 'number' ? raw : parseFloat(String(raw).replace(',','.'));
      if (isNaN(val)) continue;
      const stored = Math.round(val / 100 * 1e6) / 1e6;
      timeseries[key][STATIC_YEAR] = stored;
      latest[key] = stored;
    }
    communes[nis] = { timeseries, latest };
  }
  const metadata = {};
  for (const c of COLS) {
    const vals = Object.values(communes).map(co => co.latest[c.key]).filter(v => v != null);
    if (!vals.length) continue;
    const min = Math.min(...vals), max = Math.max(...vals);
    metadata[c.key] = { label:c.label, labelEn:c.labelEn, unit:c.unit,
      years:[STATIC_YEAR], latestYear:STATIC_YEAR,
      min:Math.floor(min*1000)/1000, max:Math.ceil(max*1000)/1000,
      step: (max-min)>0.1?0.001:0.0001, count:Object.keys(communes).length };
  }
  return { communes, metadata, allYears:[STATIC_YEAR] };
}

// =====================================================================
// 6. Excel (families.xlsx) → family structure time series
// =====================================================================

const FAMILY_BANDS = [
  { start: 2,   key: 'fam_single_parent',     label: 'Aandeel alleenstaande ouders',            labelEn: 'Share single parents',           unit: '%' },
  { start: 27,  key: 'fam_alone_under30',      label: 'Aandeel alleenwonenden jonger dan 30',    labelEn: 'Share living alone <30',         unit: '%' },
  { start: 52,  key: 'fam_alone_18_29',        label: 'Alleenwonenden 18-29 (% van leeftijdsgroep)', labelEn: 'Living alone 18-29 (% of group)', unit: '%' },
  { start: 77,  key: 'fam_alone_65plus_pct',   label: 'Alleenwonenden 65+ (% van 65+)',          labelEn: 'Living alone 65+ (% of 65+)',    unit: '%' },
  { start: 102, key: 'fam_alone_65plus_hh',    label: 'Alleenwonenden 65+ (% huishoudens)',      labelEn: 'Lone 65+ households (%)',        unit: '%' },
  { start: 127, key: 'fam_couple_children',    label: 'Koppels met kinderen',                   labelEn: 'Couples with children',          unit: '%' },
  { start: 152, key: 'fam_couple_no_children', label: 'Koppels zonder kinderen',                labelEn: 'Couples without children',       unit: '%' },
];

function convertFamilies() {
  console.log("\nParsing families.xlsx...");
  const xlsxFile = existsSync(join(ROOT, 'data', 'families.xlsx')) ? join(ROOT, 'data', 'families.xlsx') : join(ROOT, 'families.xlsx');
  const wb = XLSX.readFile(xlsxFile);
  const ws = wb.Sheets[wb.SheetNames[0]];

  function cell(r, c) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cl = ws[addr];
    return cl ? cl.v : null;
  }

  // Build colMap (25 years: 2001–2025, year in row 2)
  const colMap = [];
  for (const band of FAMILY_BANDS) {
    for (let i = 0; i < 25; i++) {
      const col = band.start + i;
      const year = cell(2, col);
      if (typeof year !== 'number') continue;
      colMap.push({ col, key: band.key, year });
    }
  }
  console.log(`  ${colMap.length} (col, key, year) entries`);

  const communes = {};

  for (let r = 7; r <= 25; r++) {
    const rawCode = cell(r, 0);
    if (rawCode === null || rawCode === undefined) continue;
    const niscode = String(rawCode).trim().padStart(5, '0');
    if (!/^\d{5}$/.test(niscode)) continue;
    if (!niscode.startsWith('21')) continue;

    const timeseries = {};
    const latest = {};
    for (const band of FAMILY_BANDS) timeseries[band.key] = {};

    for (const { col, key, year } of colMap) {
      const raw = cell(r, col);
      if (raw === null || raw === undefined) continue;
      const val = typeof raw === 'number' ? raw : parseFloat(String(raw).replace(',', '.'));
      if (isNaN(val)) continue;
      timeseries[key][year] = Math.round(val / 100 * 1e6) / 1e6;
    }

    for (const band of FAMILY_BANDS) {
      const sortedYears = Object.keys(timeseries[band.key]).map(Number).sort((a, b) => b - a);
      if (sortedYears.length > 0) latest[band.key] = timeseries[band.key][sortedYears[0]];
    }

    communes[niscode] = { timeseries, latest };
  }

  console.log(`  Read ${Object.keys(communes).length} communes`);

  const metadata = {};
  const allYearsSet = new Set();

  for (const band of FAMILY_BANDS) {
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
    const sortedYears = [...yearsWithData].sort((a, b) => a - b);

    metadata[band.key] = {
      label: band.label, labelEn: band.labelEn, unit: band.unit,
      years: sortedYears, latestYear: sortedYears[sortedYears.length - 1],
      min: Math.floor(min * 1000) / 1000,
      max: Math.ceil(max * 1000) / 1000,
      step: 0.001, count: Object.keys(communes).length
    };
  }

  const allYears = [...allYearsSet].sort((a, b) => a - b);
  console.log(`  Family year range: ${allYears[0]}–${allYears[allYears.length - 1]}`);
  console.log(`  Family indicators: ${Object.keys(metadata).length}`);

  return { communes, metadata, allYears };
}

// =====================================================================
// 7. Excel (migratieachtergrond.xlsx) → migration background (static 2021)
// =====================================================================

// No NIS codes in source — match by Dutch commune name (case-insensitive, trimmed)
const NISCODE_BY_NAME_DUT = {
  "anderlecht":                "21001",
  "oudergem":                  "21002",
  "sint-agatha-berchem":       "21003",
  "brussel":                   "21004",
  "etterbeek":                 "21005",
  "evere":                     "21006",
  "vorst":                     "21007",
  "ganshoren":                 "21008",
  "elsene":                    "21009",
  "jette":                     "21010",
  "koekelberg":                "21011",
  "sint-jans-molenbeek":       "21012",
  "sint-gillis":               "21013",
  "sint-joost-ten-node":       "21014",
  "schaarbeek":                "21015",
  "ukkel":                     "21016",
  "watermaal-bosvoorde":       "21017",
  "sint-lambrechts-woluwe":    "21018",
  "sint-pieters-woluwe":       "21019",
};

const MIGBG_YEAR = 2021;

const MIGBG_COLS = [
  { col: 1, key: 'migbg_total',         label: 'Migratie-achtergrond',             labelEn: 'Migration background',                  unit: '%' },
  { col: 2, key: 'migbg_belgian_with',  label: 'Belgen met migratieachtergrond',   labelEn: 'Belgians with migration background',    unit: '%' },
  { col: 3, key: 'migbg_belgian_without', label: 'Belgen zonder migratieachtergrond', labelEn: 'Belgians without migration background', unit: '%' },
];

function convertMigBg() {
  console.log("\nParsing migratieachtergrond.xlsx...");
  const xlsxFile = existsSync(join(ROOT, 'data', 'migratieachtergrond.xlsx')) ? join(ROOT, 'data', 'migratieachtergrond.xlsx') : join(ROOT, 'migratieachtergrond.xlsx');
  const wb = XLSX.readFile(xlsxFile);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const range = XLSX.utils.decode_range(ws["!ref"]);

  function cell(r, c) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cl = ws[addr];
    return cl ? cl.v : null;
  }

  const communes = {};
  let unmatched = [];

  for (let r = 1; r <= range.e.r; r++) {
    const rawName = cell(r, 0);
    if (rawName === null || rawName === undefined) continue;
    const key = String(rawName).trim().toLowerCase();
    const niscode = NISCODE_BY_NAME_DUT[key];
    if (!niscode) { unmatched.push(String(rawName).trim()); continue; }

    const timeseries = {};
    const latest = {};
    for (const c of MIGBG_COLS) timeseries[c.key] = {};

    for (const { col, key: varKey } of MIGBG_COLS) {
      const val = cell(r, col);
      if (val === null || val === undefined) continue;
      const num = typeof val === 'number' ? val : parseFloat(String(val).replace(',', '.'));
      if (isNaN(num)) continue;
      // Values are already fractions (0.889 = 88.9%)
      timeseries[varKey][MIGBG_YEAR] = num;
      latest[varKey] = num;
    }

    communes[niscode] = { timeseries, latest };
  }

  console.log(`  Read ${Object.keys(communes).length} communes`);
  if (unmatched.length) console.warn(`  Unmatched names: ${unmatched.join(', ')}`);

  const metadata = {};
  for (const c of MIGBG_COLS) {
    const vals = Object.values(communes).map(co => co.latest[c.key]).filter(v => v != null);
    if (vals.length === 0) continue;
    const min = Math.min(...vals);
    const max = Math.max(...vals);
    metadata[c.key] = {
      label: c.label, labelEn: c.labelEn, unit: c.unit,
      years: [MIGBG_YEAR], latestYear: MIGBG_YEAR,
      min: Math.floor(min * 1000) / 1000,
      max: Math.ceil(max * 1000) / 1000,
      step: 0.001, count: Object.keys(communes).length
    };
  }

  console.log(`  Migration background indicators: ${Object.keys(metadata).length}`);
  return { communes, metadata, allYears: [MIGBG_YEAR] };
}

// =====================================================================
// 8. Excel (share of Belgians per commune.xlsx) → nationality time series
// =====================================================================

const NATIONALITY_BANDS = [
  { start: 2,  key: 'nat_eu14_pct',    label: 'Aandeel EU14-onderdanen',  labelEn: 'Share EU14 nationals',  unit: '%' },
  { start: 28, key: 'nat_belgian_pct', label: 'Aandeel Belgen',           labelEn: 'Share Belgians',        unit: '%' },
];

function convertNationality() {
  console.log("\nParsing share of Belgians per commune.xlsx...");
  const xlsxFile = existsSync(join(ROOT, 'data', 'share of Belgians per commune.xlsx')) ? join(ROOT, 'data', 'share of Belgians per commune.xlsx') : join(ROOT, 'share of Belgians per commune.xlsx');
  const wb = XLSX.readFile(xlsxFile);
  const ws = wb.Sheets[wb.SheetNames[0]];

  function cell(r, c) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cl = ws[addr];
    return cl ? cl.v : null;
  }

  // Build colMap (26 years: 2000–2025, year in row 2)
  const colMap = [];
  for (const band of NATIONALITY_BANDS) {
    for (let i = 0; i < 26; i++) {
      const col = band.start + i;
      const year = cell(2, col);
      if (typeof year !== 'number') continue;
      colMap.push({ col, key: band.key, year });
    }
  }
  console.log(`  ${colMap.length} (col, key, year) entries`);

  const communes = {};

  // Data rows 3–21 (19 communes, no region total row)
  for (let r = 3; r <= 21; r++) {
    const rawCode = cell(r, 0);
    if (rawCode === null || rawCode === undefined) continue;
    const niscode = String(rawCode).trim().padStart(5, '0');
    if (!/^\d{5}$/.test(niscode)) continue;
    if (!niscode.startsWith('21')) continue;

    const timeseries = {};
    const latest = {};
    for (const band of NATIONALITY_BANDS) timeseries[band.key] = {};

    for (const { col, key, year } of colMap) {
      const raw = cell(r, col);
      if (raw === null || raw === undefined) continue;
      const val = typeof raw === 'number' ? raw : parseFloat(String(raw).replace(',', '.'));
      if (isNaN(val)) continue;
      timeseries[key][year] = Math.round(val / 100 * 1e6) / 1e6;
    }

    for (const band of NATIONALITY_BANDS) {
      const sortedYears = Object.keys(timeseries[band.key]).map(Number).sort((a, b) => b - a);
      if (sortedYears.length > 0) latest[band.key] = timeseries[band.key][sortedYears[0]];
    }

    communes[niscode] = { timeseries, latest };
  }

  console.log(`  Read ${Object.keys(communes).length} communes`);

  const metadata = {};
  const allYearsSet = new Set();

  for (const band of NATIONALITY_BANDS) {
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
    const sortedYears = [...yearsWithData].sort((a, b) => a - b);

    metadata[band.key] = {
      label: band.label, labelEn: band.labelEn, unit: band.unit,
      years: sortedYears, latestYear: sortedYears[sortedYears.length - 1],
      min: Math.floor(min * 1000) / 1000,
      max: Math.ceil(max * 1000) / 1000,
      step: 0.001, count: Object.keys(communes).length
    };
  }

  const allYears = [...allYearsSet].sort((a, b) => a - b);
  console.log(`  Nationality year range: ${allYears[0]}–${allYears[allYears.length - 1]}`);
  console.log(`  Nationality indicators: ${Object.keys(metadata).length}`);

  return { communes, metadata, allYears };
}

// =====================================================================
// 9. Excel (pop1992-mov_fr.xlsx) → population movement 1992–2024
//    New variables: births, deaths, natural_balance,
//                   intl_in_num, intl_out_num, naturalizations
//    Extended time series (→ 1992): total_population, internal_in_num,
//    internal_out_num, internal_balance_num, international_balance_num,
//    plus computed pct variants
// =====================================================================

const POP_MOV_INDICATORS = [
  // col → variable key mapping (per row in each year-sheet)
  { col: 2,  key: 'total_population',          newVar: false },
  { col: 3,  key: 'births',                    newVar: true,
    label: 'Geboorten', labelEn: 'Births', unit: 'personen' },
  { col: 4,  key: 'deaths',                    newVar: true,
    label: 'Overlijdens', labelEn: 'Deaths', unit: 'personen' },
  { col: 5,  key: 'natural_balance',           newVar: true,
    label: 'Natuurlijk saldo', labelEn: 'Natural balance (births − deaths)', unit: 'personen' },
  { col: 6,  key: 'internal_in_num',           newVar: false },
  { col: 7,  key: 'internal_out_num',          newVar: false },
  { col: 8,  key: 'internal_balance_num',      newVar: false },
  { col: 9,  key: 'intl_in_num',               newVar: true,
    label: 'Internationale immigratie', labelEn: 'International immigration (gross in)', unit: 'personen' },
  { col: 12, key: 'intl_out_num',              newVar: true,
    label: 'Internationale emigratie', labelEn: 'International emigration (gross out)', unit: 'personen' },
  { col: 15, key: 'international_balance_num', newVar: false },
  { col: 21, key: 'naturalizations',           newVar: true,
    label: 'Naturalisaties', labelEn: 'Naturalizations (non-Belgian → Belgian)', unit: 'personen' },
];

// Percentage variants computed from absolute numbers + population
const PCT_COMPUTED = [
  { numKey: 'internal_in_num',           pctKey: 'internal_in_pct' },
  { numKey: 'internal_out_num',          pctKey: 'internal_out_pct' },
  { numKey: 'internal_balance_num',      pctKey: 'internal_balance_pct' },
  { numKey: 'international_balance_num', pctKey: 'international_balance_pct' },
];

function convertPopMovement() {
  const possiblePaths = [
    join(ROOT, 'data', 'pop1992-mov_fr.xlsx'),
    join(ROOT, 'pop1992-mov_fr.xlsx'),
    'C:/Users/sushe/Downloads/Datasets/pop1992-mov_fr.xlsx',
  ];

  let xlsxPath = null;
  for (const p of possiblePaths) {
    if (existsSync(p)) { xlsxPath = p; break; }
  }

  if (!xlsxPath) {
    console.log('\nSkipping pop1992-mov_fr.xlsx — file not found');
    return null;
  }

  console.log(`\nParsing ${xlsxPath}...`);
  const wb = XLSX.readFile(xlsxPath);
  const years = wb.SheetNames
    .map(n => parseInt(n, 10))
    .filter(n => !isNaN(n))
    .sort((a, b) => a - b);
  console.log(`  Found ${years.length} year sheets: ${years[0]}–${years[years.length - 1]}`);

  const communes = {};

  for (const year of years) {
    const ws = wb.Sheets[String(year)];
    if (!ws) continue;
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

    for (const row of rows) {
      const rawCode = row[0];
      if (!rawCode) continue;
      const niscode = String(rawCode).trim().padStart(5, '0');
      if (!niscode.startsWith('21') || niscode === '21000') continue;
      if (!/^21\d{3}$/.test(niscode)) continue;

      if (!communes[niscode]) {
        communes[niscode] = { timeseries: {}, latest: {} };
        for (const ind of POP_MOV_INDICATORS) communes[niscode].timeseries[ind.key] = {};
        for (const pc of PCT_COMPUTED) communes[niscode].timeseries[pc.pctKey] = {};
      }

      const pop = typeof row[2] === 'number' ? row[2] : null;

      for (const ind of POP_MOV_INDICATORS) {
        const val = row[ind.col];
        if (typeof val !== 'number') continue;
        communes[niscode].timeseries[ind.key][year] = val;
      }

      // Compute pct variants (stored as decimal fraction, e.g. 0.05 = 5%)
      if (pop && pop > 0) {
        for (const pc of PCT_COMPUTED) {
          const num = communes[niscode].timeseries[pc.numKey][year];
          if (num !== undefined) {
            communes[niscode].timeseries[pc.pctKey][year] =
              Math.round((num / pop) * 1e6) / 1e6;
          }
        }
      }
    }
  }

  // Compute latest values
  for (const c of Object.values(communes)) {
    for (const [key, ts] of Object.entries(c.timeseries)) {
      const sortedYears = Object.keys(ts).map(Number).sort((a, b) => b - a);
      if (sortedYears.length > 0) c.latest[key] = ts[sortedYears[0]];
    }
  }

  console.log(`  Read ${Object.keys(communes).length} communes`);

  // Build metadata only for NEW variables (existing ones keep their metadata;
  // years will be updated by mergeDatasetInto)
  const newIndicators = POP_MOV_INDICATORS.filter(i => i.newVar);
  const metadata = {};
  const allYearsSet = new Set();

  for (const ind of [...newIndicators]) {
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
    const step = range_ > 10000 ? 100 : range_ > 1000 ? 50 : range_ > 100 ? 5 : 1;
    const sortedYears = [...yearsWithData].sort((a, b) => a - b);

    metadata[ind.key] = {
      label: ind.label, labelEn: ind.labelEn, unit: ind.unit,
      years: sortedYears, latestYear: sortedYears[sortedYears.length - 1],
      min: Math.floor(min), max: Math.ceil(max),
      step, count: Object.keys(communes).length,
    };
  }

  // Also collect years from all variables for allYears
  for (const c of Object.values(communes)) {
    for (const ts of Object.values(c.timeseries)) {
      for (const yr of Object.keys(ts)) allYearsSet.add(Number(yr));
    }
  }

  const allYears = [...allYearsSet].sort((a, b) => a - b);
  console.log(`  Year range: ${allYears[0]}–${allYears[allYears.length - 1]}`);
  console.log(`  New indicators: ${Object.keys(metadata).length}`);

  return { communes, metadata, allYears };
}

// =====================================================================
// mergeDatasetInto — year-by-year merge (adds missing years without
// overwriting existing values; updates metadata years arrays)
// =====================================================================
function mergeDatasetInto(main, extra, label) {
  if (!extra) return;
  console.log(`\nMerging ${label} (extend mode)...`);
  let merged = 0;

  for (const [niscode, extraC] of Object.entries(extra.communes)) {
    const mainC = main.communes[niscode];
    if (!mainC) continue;

    for (const [key, ts] of Object.entries(extraC.timeseries)) {
      if (!mainC.timeseries[key]) mainC.timeseries[key] = {};
      // Add only years not already present
      for (const [yr, val] of Object.entries(ts)) {
        if (mainC.timeseries[key][yr] === undefined) {
          mainC.timeseries[key][yr] = val;
        }
      }
      // Recompute latest from merged timeseries
      const sortedYears = Object.keys(mainC.timeseries[key]).map(Number).sort((a, b) => b - a);
      if (sortedYears.length > 0) mainC.latest[key] = mainC.timeseries[key][sortedYears[0]];
    }
    merged++;
  }

  // Merge metadata: new keys → add fully; existing keys → extend years array
  for (const [key, meta] of Object.entries(extra.metadata)) {
    if (main.metadata[key]) {
      const yearSet = new Set([...main.metadata[key].years, ...meta.years]);
      main.metadata[key].years = [...yearSet].sort((a, b) => a - b);
    } else {
      main.metadata[key] = meta;
    }
  }

  // Also extend years for existing variables not in extra.metadata
  // (e.g. total_population — metadata stays, but timeseries now has 1992+ data)
  for (const ind of POP_MOV_INDICATORS) {
    if (main.metadata[ind.key]) {
      const yearSet = new Set(main.metadata[ind.key].years);
      for (const c of Object.values(extra.communes)) {
        for (const yr of Object.keys(c.timeseries[ind.key] || {})) yearSet.add(Number(yr));
      }
      main.metadata[ind.key].years = [...yearSet].sort((a, b) => a - b);
    }
  }
  // Same for pct variants
  for (const pc of PCT_COMPUTED) {
    if (main.metadata[pc.pctKey]) {
      const yearSet = new Set(main.metadata[pc.pctKey].years);
      for (const c of Object.values(extra.communes)) {
        for (const yr of Object.keys(c.timeseries[pc.pctKey] || {})) yearSet.add(Number(yr));
      }
      main.metadata[pc.pctKey].years = [...yearSet].sort((a, b) => a - b);
    }
  }

  const yearSet = new Set(main.allYears);
  for (const yr of extra.allYears) yearSet.add(yr);
  main.allYears = [...yearSet].sort((a, b) => a - b);

  console.log(`  Extended ${merged} communes`);
}

// =====================================================================
// 10. Excel (PRI_RC.xlsx) → property tax & cadastral value indicators
//     Source: SPF Finances / Statbel 2024
//     Indicators:
//       pri_avg_dwelling  — avg annual property tax per dwelling (€)
//       pri_avg_parcel    — avg annual property tax per parcel (€)
//       pri_std_dev       — std deviation of property tax (€) = internal inequality
//       rc_per_ha         — taxable cadastral income per hectare (€/ha)
//       rc_total          — total taxable cadastral income (€)
//       rc_pct_low        — share of dwellings with RC < 600 (low-value housing, %)
//       rc_pct_high       — share of dwellings with RC ≥ 2000 (high-value housing, %)
// =====================================================================

const NISCODE_BY_NAME_FR_PRI = {
  'anderlecht':              '21001',
  'auderghem':               '21002',
  'berchem-sainte-agathe':   '21003',
  'bruxelles':               '21004',
  'etterbeek':               '21005',
  'evere':                   '21006',
  'forest':                  '21007',
  'ganshoren':               '21008',
  'ixelles':                 '21009',
  'jette':                   '21010',
  'koekelberg':              '21011',
  'molenbeek-saint-jean':    '21012',
  'saint-gilles':            '21013',
  'saint-josse-ten-noode':   '21014',
  'schaerbeek':              '21015',
  'uccle':                   '21016',
  'watermael-boitsfort':     '21017',
  'woluwe-saint-lambert':    '21018',
  'woluwe-saint-pierre':     '21019',
};

const PRI_RC_YEAR = 2024;

const PRI_RC_COLS = [
  { key: 'pri_avg_dwelling', label: 'Gem. onroerende voorheffing/woning',    labelEn: 'Avg property tax per dwelling',   unit: '€' },
  { key: 'pri_avg_parcel',   label: 'Gem. onroerende voorheffing/perceel',   labelEn: 'Avg property tax per parcel',     unit: '€' },
  { key: 'pri_std_dev',      label: 'Spreiding onroerende voorheffing',      labelEn: 'Property tax spread (std dev)',   unit: '€' },
  { key: 'rc_per_ha',        label: 'Kadastraal inkomen per hectare',        labelEn: 'Cadastral income per hectare',    unit: '€/ha' },
  { key: 'rc_total',         label: 'Totaal belastbaar kadastraal inkomen',  labelEn: 'Total taxable cadastral income',  unit: '€' },
  { key: 'rc_pct_low',       label: 'Aandeel woningen lage KI (< 600)',      labelEn: 'Share low-value dwellings (RC < 600)',  unit: '%' },
  { key: 'rc_pct_high',      label: 'Aandeel woningen hoge KI (≥ 2000)',     labelEn: 'Share high-value dwellings (RC ≥ 2000)', unit: '%' },
];

function convertPriRc() {
  const possiblePaths = [
    join(ROOT, 'data', 'PRI_RC.xlsx'),
    join(ROOT, 'PRI_RC.xlsx'),
    'C:/Users/sushe/Downloads/Datasets/PRI_RC.xlsx',
  ];
  let xlsxPath = null;
  for (const p of possiblePaths) { if (existsSync(p)) { xlsxPath = p; break; } }
  if (!xlsxPath) { console.log('\nSkipping PRI_RC.xlsx — file not found'); return null; }

  console.log(`\nParsing ${xlsxPath}...`);
  const wb = XLSX.readFile(xlsxPath);

  // --- Dispersion sheet: pri_avg_dwelling, pri_avg_parcel, pri_std_dev ---
  const dispRows = XLSX.utils.sheet_to_json(wb.Sheets['Dispersion'], { header: 1 }).slice(1);
  const dispByNis = {};
  for (const row of dispRows) {
    const name = String(row[0] || '').trim().toLowerCase();
    const nis = NISCODE_BY_NAME_FR_PRI[name];
    if (!nis) continue;
    const stdDev   = typeof row[1] === 'number' ? row[1] : null;
    const avgParcel = typeof row[2] === 'number' ? row[2] : null;
    const avgDwell  = typeof row[3] === 'number' ? row[3] : null;
    dispByNis[nis] = { stdDev, avgParcel, avgDwell };
  }

  // --- Statbel sheet: rc_total, rc_per_ha ---
  const statRows = XLSX.utils.sheet_to_json(wb.Sheets['Statbel'], { header: 1 }).slice(1);
  const statByNis = {};
  for (const row of statRows) {
    const name = String(row[0] || '').trim().toLowerCase();
    const nis = NISCODE_BY_NAME_FR_PRI[name];
    if (!nis) continue;
    const rcTotal = typeof row[1] === 'number' ? row[1] : null;
    const rcPerHa = typeof row[3] === 'number' ? row[3] : null;
    statByNis[nis] = { rcTotal, rcPerHa };
  }

  // --- Communes sheet: rc_pct_low, rc_pct_high ---
  const comRows = XLSX.utils.sheet_to_json(wb.Sheets['Communes'], { header: 1 }).slice(1);
  const comByNis = {};

  function bandLower(rc) {
    if (rc === 0) return 0;
    if (typeof rc === 'string') { const m = rc.match(/^(\d+)/); return m ? parseInt(m[1]) : null; }
    return null;
  }

  for (const row of comRows) {
    const name = String(row[0] || '').trim().toLowerCase();
    const nis = NISCODE_BY_NAME_FR_PRI[name];
    if (!nis) continue;
    const dwellings = typeof row[3] === 'number' ? row[3] : 0;
    const lower = bandLower(row[1]);
    if (lower === null) continue;
    if (!comByNis[nis]) comByNis[nis] = { total: 0, low: 0, high: 0 };
    comByNis[nis].total  += dwellings;
    if (lower < 600)   comByNis[nis].low  += dwellings;
    if (lower >= 2000) comByNis[nis].high += dwellings;
  }

  // --- Assemble communes ---
  const communes = {};
  const allNis = new Set([...Object.keys(dispByNis), ...Object.keys(statByNis), ...Object.keys(comByNis)]);

  for (const nis of allNis) {
    const d = dispByNis[nis] || {};
    const s = statByNis[nis] || {};
    const c = comByNis[nis]  || {};

    const rcPctLow  = c.total > 0 ? Math.round(c.low  / c.total * 1e6) / 1e6 : null;
    const rcPctHigh = c.total > 0 ? Math.round(c.high / c.total * 1e6) / 1e6 : null;

    const vals = {
      pri_avg_dwelling: d.avgDwell  != null ? Math.round(d.avgDwell)  : null,
      pri_avg_parcel:   d.avgParcel != null ? Math.round(d.avgParcel) : null,
      pri_std_dev:      d.stdDev    != null ? Math.round(d.stdDev)    : null,
      rc_per_ha:        s.rcPerHa   != null ? Math.round(s.rcPerHa)   : null,
      rc_total:         s.rcTotal   != null ? Math.round(s.rcTotal)   : null,
      rc_pct_low:       rcPctLow,
      rc_pct_high:      rcPctHigh,
    };

    const timeseries = {};
    const latest = {};
    for (const col of PRI_RC_COLS) {
      timeseries[col.key] = {};
      if (vals[col.key] != null) {
        timeseries[col.key][PRI_RC_YEAR] = vals[col.key];
        latest[col.key] = vals[col.key];
      }
    }
    communes[nis] = { timeseries, latest };
  }

  console.log(`  Read ${Object.keys(communes).length} communes`);

  // --- Metadata ---
  const metadata = {};
  for (const col of PRI_RC_COLS) {
    const vals = Object.values(communes).map(c => c.latest[col.key]).filter(v => v != null);
    if (!vals.length) continue;
    const min = Math.min(...vals);
    const max = Math.max(...vals);
    const range_ = max - min;
    let step;
    if (col.unit === '%') step = 0.001;
    else if (range_ > 100000) step = 1000;
    else if (range_ > 10000) step = 100;
    else step = 1;
    metadata[col.key] = {
      label: col.label, labelEn: col.labelEn, unit: col.unit,
      years: [PRI_RC_YEAR], latestYear: PRI_RC_YEAR,
      min: col.unit === '%' ? Math.floor(min * 1000) / 1000 : Math.floor(min),
      max: col.unit === '%' ? Math.ceil(max  * 1000) / 1000 : Math.ceil(max),
      step, count: Object.keys(communes).length,
    };
  }

  console.log(`  Property indicators: ${Object.keys(metadata).length}`);
  return { communes, metadata, allYears: [PRI_RC_YEAR] };
}

// =====================================================================
// 11. CSV (SPF Finances Concentration, annual 2011–2025) → property wealth time series
//     Source: finances.belgium.be — Characteristics of cadastral parcels — Concentration of cadastral income
//     Files: MunicipalityWideConcentration_YYYYMMDD.csv (one per year)
//     Overrides single-year PRI_RC values for: pri_avg_dwelling, rc_pct_low, rc_pct_high
// =====================================================================

function convertCadastralIncome() {
  const ciDirs = [
    join(ROOT, 'data', 'Cadastral Income'),
    join(ROOT, 'Cadastral Income'),
    'C:/Users/sushe/Downloads/Datasets/Cadastral Income',
  ];
  let ciDir = null;
  for (const d of ciDirs) { if (existsSync(d)) { ciDir = d; break; } }
  if (!ciDir) { console.log('\nSkipping Cadastral Income CSVs — folder not found'); return null; }
  console.log(`\nParsing Cadastral Income CSVs from ${ciDir}...`);

  const YEARS = [2011,2012,2013,2014,2015,2016,2017,2018,2019,2020,2021,2022,2023,2024,2025];
  const commData = {};

  for (const year of YEARS) {
    const csvPath = join(ciDir, String(year), `MunicipalityWideConcentration_${year}0101.csv`);
    if (!existsSync(csvPath)) { console.log(`  Skipping ${year} — file not found`); continue; }

    const raw   = readFileSync(csvPath, 'utf8').replace(/^\uFEFF/, '');
    const lines = raw.split('\n');
    const hdr   = lines[0].split(';');
    const hi    = c => hdr.indexOf(c);

    // Group by NISCode for Brussels only
    const byNis = {};
    for (const line of lines.slice(1)) {
      if (!line.trim()) continue;
      const c = line.split(';');
      const nis = c[0];
      if (!nis.startsWith('21')) continue;
      if (!byNis[nis]) byNis[nis] = [];
      byNis[nis].push(c);
    }

    for (const [nis, rows] of Object.entries(byNis)) {
      let totalRC = 0, totalHousings = 0, lowCount = 0, highCount = 0;

      for (const c of rows) {
        const rangeName  = c[hi('Range')] || '';
        const rangeLow   = parseInt(rangeName.replace('Range', '')) || 0;
        const housings   = parseInt(c[hi('HousingsNumber')]) || 0;
        const rc         = parseFloat(c[hi('TotalCadastralIncome')]) || 0;

        totalRC       += rc;
        totalHousings += housings;
        if (rangeLow < 600)   lowCount  += housings;
        if (rangeLow >= 2000) highCount += housings;
      }

      if (totalHousings === 0) continue;
      const pct = (n, d) => d > 0 ? Math.round(n / d * 1e6) / 1e6 : null;

      if (!commData[nis]) commData[nis] = {
        pri_avg_dwelling: {}, rc_pct_low: {}, rc_pct_high: {}
      };
      commData[nis].pri_avg_dwelling[year] = Math.round(totalRC / totalHousings);
      commData[nis].rc_pct_low[year]       = pct(lowCount,  totalHousings);
      commData[nis].rc_pct_high[year]      = pct(highCount, totalHousings);
    }
    console.log(`  ${year}: processed ${Object.keys(byNis).length} Brussels communes`);
  }

  // Build output in mergeDataset-compatible format
  const VARS = [
    { key: 'pri_avg_dwelling', label: 'Gem. onroerende voorheffing/woning', labelEn: 'Avg property tax per dwelling', unit: '€' },
    { key: 'rc_pct_low',       label: 'Aandeel woningen lage KI (< 600)',   labelEn: 'Share low-value dwellings (RC < 600)',   unit: '%' },
    { key: 'rc_pct_high',      label: 'Aandeel woningen hoge KI (≥ 2000)',  labelEn: 'Share high-value dwellings (RC ≥ 2000)', unit: '%' },
  ];

  const allYears = YEARS.filter(y =>
    Object.values(commData).some(c => c.pri_avg_dwelling[y] != null)
  );

  const communes = {};
  for (const [nis, cd] of Object.entries(commData)) {
    const timeseries = { pri_avg_dwelling: cd.pri_avg_dwelling, rc_pct_low: cd.rc_pct_low, rc_pct_high: cd.rc_pct_high };
    const latest = {};
    for (const v of VARS) {
      const yrs = Object.keys(timeseries[v.key]).map(Number).sort((a,b) => b-a);
      if (yrs.length) latest[v.key] = timeseries[v.key][yrs[0]];
    }
    communes[nis] = { timeseries, latest };
  }

  const metadata = {};
  for (const v of VARS) {
    const vals = Object.values(communes).map(c => c.latest[v.key]).filter(x => x != null);
    if (!vals.length) continue;
    const min = Math.min(...vals), max = Math.max(...vals);
    const range_ = max - min;
    const step = v.unit === '%' ? 0.001 : (range_ > 10000 ? 100 : range_ > 1000 ? 50 : 1);
    metadata[v.key] = {
      label: v.label, labelEn: v.labelEn, unit: v.unit,
      years: allYears, latestYear: allYears[allYears.length - 1],
      min: v.unit === '%' ? Math.floor(min*1000)/1000 : Math.floor(min),
      max: v.unit === '%' ? Math.ceil(max*1000)/1000  : Math.ceil(max),
      step, count: Object.keys(communes).length,
    };
  }

  console.log(`  Cadastral income: ${allYears.length} years (${allYears[0]}–${allYears[allYears.length-1]}), ${Object.keys(metadata).length} indicators`);
  return { communes, metadata, allYears };
}

// =====================================================================
// 12. Excel (TF_PSNL_INC_TAX_MUNTY.xlsx) → personal income tax 2005–2023
//     Source: Statbel — one row per commune per year, all Belgian municipalities
//     Derived indicators (computed from raw totals):
//       avg_net_income    — avg net income per tax declaration (€)
//       zero_income_rate  — share of filings with zero declared income (%)
//       avg_prof_income   — avg professional/wage income per declaration (€)
//       tax_per_resident  — total income taxes collected per resident (€)
// =====================================================================

const INCOME_TAX_INDICATORS = [
  { key: 'avg_net_income',   label: 'Gemiddeld netto-inkomen',         labelEn: 'Avg net income per declaration',     unit: '€' },
  { key: 'zero_income_rate', label: 'Aandeel nul-inkomens',            labelEn: 'Zero income rate',                   unit: '%' },
  { key: 'avg_prof_income',  label: 'Gemiddeld beroepsinkomen',        labelEn: 'Avg professional income',            unit: '€' },
  { key: 'tax_per_resident', label: 'Belasting per inwoner',           labelEn: 'Total income tax per resident',      unit: '€' },
];

function convertIncomeTax() {
  const possiblePaths = [
    join(ROOT, 'data', 'TF_PSNL_INC_TAX_MUNTY.xlsx'),
    join(ROOT, 'TF_PSNL_INC_TAX_MUNTY.xlsx'),
    'C:/Users/sushe/Downloads/Datasets/TF_PSNL_INC_TAX_MUNTY.xlsx',
  ];
  let xlsxPath = null;
  for (const p of possiblePaths) { if (existsSync(p)) { xlsxPath = p; break; } }
  if (!xlsxPath) { console.log('\nSkipping TF_PSNL_INC_TAX_MUNTY.xlsx — file not found'); return null; }

  console.log(`\nParsing ${xlsxPath}...`);
  const wb   = XLSX.readFile(xlsxPath);
  const rows = XLSX.utils.sheet_to_json(wb.Sheets['TF_PSNL_INC_TAX_MUNTY'], { header: 1 }).slice(1);
  const bxl  = rows.filter(r => String(r[1]).startsWith('21'));

  const communes = {};

  for (const r of bxl) {
    const year    = r[0];
    const niscode = String(r[1]).trim().padStart(5, '0');
    if (!niscode.startsWith('21') || niscode === '21000') continue;

    if (!communes[niscode]) {
      communes[niscode] = { timeseries: {}, latest: {} };
      for (const ind of INCOME_TAX_INDICATORS) communes[niscode].timeseries[ind.key] = {};
    }

    const nbrNonZero = r[2] || 0;
    const nbrZero    = r[3] || 0;
    const totalDecl  = nbrNonZero + nbrZero;

    // avg_net_income: total net income / number of declarations
    if (r[5] && r[6]) {
      communes[niscode].timeseries['avg_net_income'][year] = Math.round(r[5] / r[6]);
    }
    // zero_income_rate: zero-income declarations / total declarations
    if (totalDecl > 0) {
      communes[niscode].timeseries['zero_income_rate'][year] =
        Math.round(nbrZero / totalDecl * 1e6) / 1e6;
    }
    // avg_prof_income: total professional income / number with professional income
    if (r[13] && r[14]) {
      communes[niscode].timeseries['avg_prof_income'][year] = Math.round(r[13] / r[14]);
    }
    // tax_per_resident: total taxes / total residents
    if (r[27] && r[29]) {
      communes[niscode].timeseries['tax_per_resident'][year] = Math.round(r[27] / r[29]);
    }
  }

  // Compute latest values
  for (const c of Object.values(communes)) {
    for (const ind of INCOME_TAX_INDICATORS) {
      const sortedYears = Object.keys(c.timeseries[ind.key]).map(Number).sort((a, b) => b - a);
      if (sortedYears.length > 0) c.latest[ind.key] = c.timeseries[ind.key][sortedYears[0]];
    }
  }

  console.log(`  Read ${Object.keys(communes).length} communes`);

  // Metadata
  const metadata = {};
  const allYearsSet = new Set();

  for (const ind of INCOME_TAX_INDICATORS) {
    const allValues = [];
    const yearsWithData = new Set();
    for (const c of Object.values(communes)) {
      for (const [yr, val] of Object.entries(c.timeseries[ind.key])) {
        allValues.push(val);
        yearsWithData.add(Number(yr));
        allYearsSet.add(Number(yr));
      }
    }
    if (!allValues.length) continue;
    const min = Math.min(...allValues);
    const max = Math.max(...allValues);
    const range_ = max - min;
    let step;
    if (ind.unit === '%') step = 0.001;
    else if (range_ > 10000) step = 100;
    else step = 1;
    const sortedYears = [...yearsWithData].sort((a, b) => a - b);
    metadata[ind.key] = {
      label: ind.label, labelEn: ind.labelEn, unit: ind.unit,
      years: sortedYears, latestYear: sortedYears[sortedYears.length - 1],
      min: ind.unit === '%' ? Math.floor(min * 1000) / 1000 : Math.floor(min),
      max: ind.unit === '%' ? Math.ceil(max  * 1000) / 1000 : Math.ceil(max),
      step, count: Object.keys(communes).length,
    };
  }

  const allYears = [...allYearsSet].sort((a, b) => a - b);
  console.log(`  Year range: ${allYears[0]}–${allYears[allYears.length - 1]}`);
  console.log(`  Income tax indicators: ${Object.keys(metadata).length}`);
  return { communes, metadata, allYears };
}

// =====================================================================
// 12. Excel (BV_tabjaar_EN.xlsx) → building permits 1996–2024
//     Source: Statbel — one row per municipality per year, long format
//     Sheet used: "Municipalities" (cols described below)
//     Col 0: Refnis (NIS code)  Col 2: Year
//     Col 4: res new dwellings  Col 5: res new apartments
//     Col 7: res new surface m² Col 8: res renovations
//     Col 9: nonres new bldgs   Col 11: nonres renovations
//     Year 2025 is present but incomplete — capped at 2024.
// =====================================================================

const BUILDING_PERMIT_INDICATORS = [
  { col: 4,  key: 'bp_res_new_dwellings', label: 'Nieuwe residentiële wooneenheden',    labelEn: 'New residential dwellings permitted',      unit: 'units' },
  { col: 5,  key: 'bp_res_apartments',    label: 'Nieuwe appartementen vergund',         labelEn: 'New apartments permitted',                 unit: 'units' },
  { col: 6,  key: 'bp_res_houses',        label: 'Nieuwe eengezinswoningen vergund',     labelEn: 'New single-family houses permitted',       unit: 'units' },
  { col: 7,  key: 'bp_res_new_surface',   label: 'Oppervlakte nieuwe woningen (m²)',    labelEn: 'New residential floor area permitted (m²)', unit: 'm²'    },
  { col: 8,  key: 'bp_res_renovations',   label: 'Renovatievergunningen residentieel',  labelEn: 'Residential renovation permits',            unit: 'units' },
  { col: 9,  key: 'bp_nonres_new',        label: 'Niet-residentiële nieuwbouw vergund', labelEn: 'Non-residential new construction permits',  unit: 'units' },
  { col: 11, key: 'bp_nonres_renov',      label: 'Renovatie niet-residentieel',         labelEn: 'Non-residential renovation permits',        unit: 'units' },
];

function convertBuildingPermits() {
  const possiblePaths = [
    join(ROOT, 'data', 'BV_tabjaar_EN.xlsx'),
    join(ROOT, 'BV_tabjaar_EN.xlsx'),
    'C:/Users/sushe/Downloads/Datasets/BV_tabjaar_EN.xlsx',
  ];
  let xlsxPath = null;
  for (const p of possiblePaths) { if (existsSync(p)) { xlsxPath = p; break; } }
  if (!xlsxPath) { console.log('\nSkipping BV_tabjaar_EN.xlsx — file not found'); return null; }

  console.log(`\nParsing ${xlsxPath}...`);
  const wb   = XLSX.readFile(xlsxPath);
  const rows = XLSX.utils.sheet_to_json(wb.Sheets['Municipalities'], { header: 1 }).slice(4);

  // Filter Brussels communes (21001–21019); skip district aggregate 21000
  const bxl = rows.filter(r => {
    const code = String(r[0]);
    return code.startsWith('21') && code !== '21000' && code.length === 5;
  });

  const communes = {};

  for (const r of bxl) {
    const niscode = String(r[0]).trim().padStart(5, '0');
    const year    = Number(r[2]);
    if (!Number.isFinite(year) || year < 1996 || year > 2024) continue; // 2025 incomplete

    if (!communes[niscode]) {
      communes[niscode] = { timeseries: {}, latest: {} };
      for (const ind of BUILDING_PERMIT_INDICATORS) communes[niscode].timeseries[ind.key] = {};
    }

    for (const ind of BUILDING_PERMIT_INDICATORS) {
      const raw = r[ind.col];
      const val = Number(raw);
      if (raw != null && Number.isFinite(val)) {
        communes[niscode].timeseries[ind.key][year] = Math.round(val);
      }
    }
  }

  // Compute latest values
  for (const c of Object.values(communes)) {
    for (const ind of BUILDING_PERMIT_INDICATORS) {
      const sortedYears = Object.keys(c.timeseries[ind.key]).map(Number).sort((a, b) => b - a);
      if (sortedYears.length > 0) c.latest[ind.key] = c.timeseries[ind.key][sortedYears[0]];
    }
  }

  // Metadata
  const metadata = {};
  const allYearsSet = new Set();

  for (const ind of BUILDING_PERMIT_INDICATORS) {
    const allValues = [];
    const yearsWithData = new Set();
    for (const c of Object.values(communes)) {
      for (const [yr, val] of Object.entries(c.timeseries[ind.key])) {
        if (val != null) {
          allValues.push(val);
          yearsWithData.add(Number(yr));
          allYearsSet.add(Number(yr));
        }
      }
    }
    if (!allValues.length) continue;
    const sortedYears = [...yearsWithData].sort((a, b) => a - b);
    metadata[ind.key] = {
      label: ind.label, labelEn: ind.labelEn, unit: ind.unit,
      years: sortedYears, latestYear: sortedYears[sortedYears.length - 1],
      min: Math.min(...allValues),
      max: Math.max(...allValues),
      step: 1,
      count: Object.keys(communes).length,
    };
  }

  const allYears = [...allYearsSet].sort((a, b) => a - b);
  console.log(`  Read ${Object.keys(communes).length} communes`);
  console.log(`  Year range: ${allYears[0]}–${allYears[allYears.length - 1]}`);
  console.log(`  Building permit indicators: ${Object.keys(metadata).length}`);
  return { communes, metadata, allYears };
}

// =====================================================================
// 13. Excel (vastgoed_2010_9999.xlsx) → real estate transaction prices 2010–2025
//     Source: Statbel — notarial property sale records
//     One sheet per year (2010–2025); long format within each sheet.
//     Filters applied: CD_niveau_refnis=5 (commune), CD_PERIOD='Y' (annual),
//     CD_REFNIS starts with '21' (Brussels).
//     Property types used:
//       'Appartementen'                    → apartments
//       'Alle huizen met 2, 3, 4 ...'      → all house types combined
//     Null suppression applies when transaction volume is too low (<5 sales).
// =====================================================================

const VASTGOED_INDICATORS = [
  // Apartment prices & volume
  { key: 're_apt_median',        typeMatch: 'Appartementen', col: 'MS_P_50_median',        label: 'Mediane prijs appartement',       labelEn: 'Median apartment sale price',       unit: '€'     },
  { key: 're_apt_p25',           typeMatch: 'Appartementen', col: 'MS_P_25',               label: '25e percentiel appartementen',    labelEn: 'Apartment price (25th percentile)', unit: '€'     },
  { key: 're_apt_transactions',  typeMatch: 'Appartementen', col: 'MS_TOTAL_TRANSACTIONS', label: 'Antal appartementtransacties',    labelEn: 'Apartment transactions',            unit: 'units' },
  // House prices & volume (all house types combined)
  { key: 're_house_median',      typeMatch: 'Alle huizen',      col: 'MS_P_50_median',        label: 'Mediane prijs huis',                    labelEn: 'Median house sale price',                      unit: '€'     },
  { key: 're_house_transactions',typeMatch: 'Alle huizen',      col: 'MS_TOTAL_TRANSACTIONS', label: 'Aantal huizentransacties',              labelEn: 'House transactions',                           unit: 'units' },
  // Terraced / semi-detached houses (2 or 3 facades)
  { key: 're_terraced_median',       typeMatch: 'Huizen met 2 of 3', col: 'MS_P_50_median',        label: 'Mediane prijs rijwoning',               labelEn: 'Median terraced/semi-det. house price',        unit: '€'     },
  { key: 're_terraced_transactions', typeMatch: 'Huizen met 2 of 3', col: 'MS_TOTAL_TRANSACTIONS', label: 'Transacties rijwoningen',               labelEn: 'Terraced/semi-det. house transactions',        unit: 'units' },
  // Detached houses (4+ facades)
  { key: 're_detached_median',       typeMatch: 'Huizen met 4 of meer', col: 'MS_P_50_median',        label: 'Mediane prijs vrijstaande woning',      labelEn: 'Median detached house price',                  unit: '€'     },
  { key: 're_detached_transactions', typeMatch: 'Huizen met 4 of meer', col: 'MS_TOTAL_TRANSACTIONS', label: 'Transacties vrijstaande woningen',      labelEn: 'Detached house transactions',                  unit: 'units' },
];

function convertVastgoed() {
  const possiblePaths = [
    join(ROOT, 'data', 'vastgoed_2010_9999.xlsx'),
    join(ROOT, 'vastgoed_2010_9999.xlsx'),
    'C:/Users/sushe/Downloads/Datasets/vastgoed_2010_9999.xlsx',
  ];
  let xlsxPath = null;
  for (const p of possiblePaths) { if (existsSync(p)) { xlsxPath = p; break; } }
  if (!xlsxPath) { console.log('\nSkipping vastgoed_2010_9999.xlsx — file not found'); return null; }

  console.log(`\nParsing ${xlsxPath}...`);
  const wb    = XLSX.readFile(xlsxPath);
  const years = wb.SheetNames.map(Number).filter(y => Number.isFinite(y) && y >= 2010 && y <= 2025).sort((a, b) => a - b);

  const communes = {};

  for (const year of years) {
    const ws = wb.Sheets[String(year)];
    if (!ws) continue;
    const rows = XLSX.utils.sheet_to_json(ws);

    // Keep only Brussels communes (level 5, annual period, NIS 21xxx ≠ 21000)
    const bxl = rows.filter(r => {
      const nis = String(r.CD_REFNIS);
      return Number(r.CD_niveau_refnis) === 5 &&
             r.CD_PERIOD === 'Y' &&
             nis.startsWith('21') &&
             nis !== '21000';
    });

    for (const r of bxl) {
      const niscode = String(r.CD_REFNIS).padStart(5, '0');
      const typeNL  = String(r.CD_TYPE_NL || '');

      if (!communes[niscode]) {
        communes[niscode] = { timeseries: {}, latest: {} };
        for (const ind of VASTGOED_INDICATORS) communes[niscode].timeseries[ind.key] = {};
      }

      for (const ind of VASTGOED_INDICATORS) {
        if (!typeNL.startsWith(ind.typeMatch)) continue;
        const val = r[ind.col];
        if (val != null && Number.isFinite(Number(val))) {
          communes[niscode].timeseries[ind.key][year] = Math.round(Number(val));
        }
      }
    }
  }

  // Compute latest values
  for (const c of Object.values(communes)) {
    for (const ind of VASTGOED_INDICATORS) {
      const sortedYears = Object.keys(c.timeseries[ind.key]).map(Number).sort((a, b) => b - a);
      if (sortedYears.length > 0) c.latest[ind.key] = c.timeseries[ind.key][sortedYears[0]];
    }
  }

  // Metadata
  const metadata   = {};
  const allYearsSet = new Set();

  for (const ind of VASTGOED_INDICATORS) {
    const allValues    = [];
    const yearsWithData = new Set();
    for (const c of Object.values(communes)) {
      for (const [yr, val] of Object.entries(c.timeseries[ind.key])) {
        if (val != null) { allValues.push(val); yearsWithData.add(Number(yr)); allYearsSet.add(Number(yr)); }
      }
    }
    if (!allValues.length) continue;
    const sortedYears = [...yearsWithData].sort((a, b) => a - b);
    metadata[ind.key] = {
      label: ind.label, labelEn: ind.labelEn, unit: ind.unit,
      years: sortedYears, latestYear: sortedYears[sortedYears.length - 1],
      min: Math.min(...allValues),
      max: Math.max(...allValues),
      step: ind.unit === '€' ? 1000 : 1,
      count: Object.keys(communes).length,
    };
  }

  const allYears = [...allYearsSet].sort((a, b) => a - b);
  console.log(`  Read ${Object.keys(communes).length} communes`);
  console.log(`  Year range: ${allYears[0]}–${allYears[allYears.length - 1]}`);
  console.log(`  Real estate indicators: ${Object.keys(metadata).length}`);
  return { communes, metadata, allYears };
}

// =====================================================================
// RUN
// =====================================================================
const geojson = await convertSHP();
const data = convertExcel();
mergeDataset(data, convertAgeExcel(),          'age data');
mergeDataset(data, convertBuildingState(),     'housing quality 2011–2025');
mergeDataset(data, convertFamilies(),          'family data');
mergeDataset(data, convertMigBg(),             'migration background');
mergeDataset(data, convertNationality(),       'nationality data');
mergeDatasetInto(data, convertPopMovement(),   'population movement 1992–2024');
mergeDataset(data, convertPriRc(),             'property tax & cadastral value');
mergeDataset(data, convertCadastralIncome(),   'cadastral income 2011–2025');   // overrides pri_avg_dwelling, rc_pct_low, rc_pct_high with time series
mergeDataset(data, convertIncomeTax(),         'personal income tax 2005–2023');
mergeDataset(data, convertBuildingPermits(),   'building permits 1996–2024');
mergeDataset(data, convertVastgoed(),          'real estate prices 2010–2025');
writeFileSync(join(ROOT, "data", "brussels-data.json"), JSON.stringify(data));
console.log(`  brussels-data.json: ${Object.keys(data.communes).length} communes, ${Object.keys(data.metadata).length} indicators`);
mergeData(geojson, data);

// ── Bundle: inject data inline into brussels-policy.html ───────────────
const mergedGeojson = JSON.parse(readFileSync(join(ROOT, "data", "brussels-merged.geojson"), 'utf8'));
const templatePath  = join(ROOT, "brussels-policy.html");
let   template      = readFileSync(templatePath, 'utf8');
const inlineScript  = `<script>\nwindow.__GEOJSON = ${JSON.stringify(mergedGeojson)};\nwindow.__DATA = ${JSON.stringify(data)};\n</script>`;
template = template.replace(
  /<!-- INLINE_DATA_START -->[\s\S]*?<!-- INLINE_DATA_END -->/,
  `<!-- INLINE_DATA_START -->\n${inlineScript}\n<!-- INLINE_DATA_END -->`
);
writeFileSync(templatePath, template, 'utf8');
console.log(`\n  Inlined data into brussels-policy.html (${(Buffer.byteLength(inlineScript,'utf8')/1024).toFixed(0)} KB added)`);

console.log("\nDone!");
