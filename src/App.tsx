import { cloneElement, isValidElement, useEffect, useMemo, useRef, useState } from "react";
import { CircleMarker, GeoJSON, MapContainer, Polygon, Polyline, Popup, TileLayer, useMap } from "react-leaflet";
import { buffer as turfBuffer } from "@turf/buffer";
import { booleanIntersects } from "@turf/boolean-intersects";
import L from "leaflet";
import { Activity, AlertTriangle, CalendarDays, ChevronDown, Download, FileSpreadsheet, Filter, Layers3, LocateFixed, MapPin, Pencil, Search, ShieldCheck, Upload, Users, X } from "lucide-react";
import * as XLSX from "xlsx";
import html2canvas from "html2canvas";
import { jsPDF } from "jspdf";
import { REGISTRIES } from "./domain/registries";

type GISLayer = { id: string; name: string; category: string; visible: boolean; data: any; source: string; importedAt: string; };
type RegistryRecord = { id: string; registryId: string; values: Record<string, string | number>; latitude: number; longitude: number; createdAt: string; };

type CaseRecord = {
  id: string;
  disease: string;
  date: string;
  hromada: string;
  settlement: string;
  latitude?: number;
  longitude?: number;
  status: string;
  age?: number;
  sex?: string;
  notes?: string;
};

const DEMO_CASES: CaseRecord[] = [
  { id: "DEMO-001", disease: "Хвороба Лайма", date: "2026-05-12", hromada: "Тернопільська міська", settlement: "Тернопіль", latitude: 49.5535, longitude: 25.5948, status: "Підтверджено", age: 42, sex: "Ж" },
  { id: "DEMO-002", disease: "Хвороба Лайма", date: "2026-05-19", hromada: "Байковецька", settlement: "Байківці", latitude: 49.5710, longitude: 25.6800, status: "Підтверджено", age: 35, sex: "Ч" },
  { id: "DEMO-003", disease: "Гостра кишкова інфекція", date: "2026-06-03", hromada: "Чортківська", settlement: "Чортків", latitude: 49.0170, longitude: 25.7980, status: "Розслідується", age: 8, sex: "Ж" },
  { id: "DEMO-004", disease: "Сказ", date: "2026-06-15", hromada: "Козівська", settlement: "Козова", latitude: 49.4350, longitude: 25.1500, status: "Підтверджено", age: 51, sex: "Ч" },
  { id: "DEMO-005", disease: "Хвороба Лайма", date: "2026-07-02", hromada: "Кременецька", settlement: "Кременець", latitude: 50.0960, longitude: 25.7240, status: "Підтверджено", age: 29, sex: "Ж" },
  { id: "DEMO-006", disease: "Сальмонельоз", date: "2026-07-08", hromada: "Тернопільська міська", settlement: "Тернопіль", latitude: 49.5600, longitude: 25.6050, status: "Розслідується", age: 16, sex: "Ч" },
  { id: "DEMO-007", disease: "Гостра кишкова інфекція", date: "2026-08-01", hromada: "Збаразька", settlement: "Збараж", latitude: 49.6630, longitude: 25.7760, status: "Підтверджено", age: 4, sex: "Ж" },
  { id: "DEMO-008", disease: "Сказ", date: "2026-08-17", hromada: "Бережанська", settlement: "Бережани", latitude: 49.4480, longitude: 24.9380, status: "Підтверджено", age: 63, sex: "Ч" },
  { id: "DEMO-009", disease: "Хвороба Лайма", date: "2026-09-04", hromada: "Теребовлянська", settlement: "Теребовля", latitude: 49.3000, longitude: 25.7050, status: "Підтверджено", age: 47, sex: "Ж" },
  { id: "DEMO-010", disease: "Гостра кишкова інфекція", date: "2026-09-21", hromada: "Заліщицька", settlement: "Заліщики", latitude: 48.6450, longitude: 25.7350, status: "Розслідується", age: 11, sex: "Ч" }
];

const DISEASES = ["Усі нозології", ...Array.from(new Set(DEMO_CASES.map((item) => item.disease)))];
const TEMPLATE_HEADERS = ["id", "disease", "date", "hromada", "settlement", "latitude", "longitude", "status", "age", "sex", "notes"];
const REGISTRY_SHEET_NAMES: Record<string, string> = { healthcare: "Healthcare", water: "Water", food: "Food", natural: "Natural", outbreaks: "Outbreaks", "bio-monitoring": "BioMonitoring", meteorology: "Meteorology", population: "Population" };
const GEO_LAYERS_SHEET = "GeoLayers";
const CASE_REGISTRY = { id: "cases", label: "Реєстр випадків", description: "Випадки інфекційних захворювань. Дані ведуться в окремому реєстрі випадків і на аркуші Cases; не змішуються з предметними реєстрами.", geometry: "Point" as const, fields: [] };

function FitToCases({ cases }: { cases: CaseRecord[] }) {
  const map = useMap();
  const first = useRef(true);
  const firstLocatedCase = cases.find((item) => Number.isFinite(item.latitude) && Number.isFinite(item.longitude));
  if (first.current && firstLocatedCase && firstLocatedCase.latitude !== undefined && firstLocatedCase.longitude !== undefined) {
    map.setView([firstLocatedCase.latitude, firstLocatedCase.longitude], 8);
    first.current = false;
  }
  return null;
}


type DrawMode = "none" | "point" | "line" | "polygon";

function FitToLayer({ layer }: { layer: GISLayer | null }) {
  const map = useMap();
  useEffect(() => {
    if (!layer?.data) return;
    try {
      const bounds = L.geoJSON(layer.data).getBounds();
      if (bounds.isValid()) map.fitBounds(bounds, { padding: [28, 28], maxZoom: 14 });
    } catch { /* Некоректна геометрія не повинна ламати карту. */ }
  }, [map, layer]);
  return null;
}

function DrawController({ mode, onMapClick }: { mode: DrawMode; onMapClick: (lat: number, lng: number) => void }) {
  const map = useMap();
  useEffect(() => {
    if (mode === "none") {
      map.getContainer().style.cursor = "";
      return;
    }
    map.getContainer().style.cursor = "crosshair";
    const handleClick = (event: L.LeafletMouseEvent) => onMapClick(event.latlng.lat, event.latlng.lng);
    map.on("click", handleClick);
    return () => {
      map.off("click", handleClick);
      map.getContainer().style.cursor = "";
    };
  }, [map, mode, onMapClick]);
  return null;
}


type FacilitySymbol = { kind: string; glyph: string; label: string };
function facilitySymbol(type: unknown): FacilitySymbol {
  const value = String(type ?? "").toLocaleLowerCase("uk-UA");
  if (/екстрен|швидк|медиц.*катастроф|103|\bемд\b|пункт базування/.test(value)) return { kind: "emergency", glyph: "✚", label: "Екстрена допомога" };
  if (/стомат|зубн/.test(value)) return { kind: "dental", glyph: "✦", label: "Стоматологія" };
  if (/лаборатор|діагност/.test(value)) return { kind: "laboratory", glyph: "⚗", label: "Лабораторія" };
  if (/амбулатор|цпмсд|первин|сімейн|фельдшер|фап|\bфп\b/.test(value)) return { kind: "primary", glyph: "⌂", label: "Первинна допомога" };
  if (/лікарн|госпітал|стаціонар/.test(value)) return { kind: "hospital", glyph: "✚", label: "Лікарня" };
  if (/приват|медичн.*центр|клінік/.test(value)) return { kind: "clinic", glyph: "●", label: "Медичний центр / клініка" };
  return { kind: "other", glyph: "•", label: "Інший медичний заклад" };
}

function FacilityMarkerLayer({ records, language }: { records: RegistryRecord[]; language: "uk" | "en" }) {
  const map = useMap();
  useEffect(() => {
    const layer = L.layerGroup().addTo(map);
    const render = () => {
      const translate = (value: string) => language === "en" ? (UI_TRANSLATIONS[value] ?? value) : value;
      layer.clearLayers();
      const size = map.getSize();
      const zoom = map.getZoom();
      const iconSize = zoom < 8 ? 17 : zoom < 11 ? 20 : zoom < 14 ? 24 : zoom < 16 ? 28 : 32;
      const visible = records.filter((record) =>
        Number.isFinite(record.latitude) && Number.isFinite(record.longitude) &&
        record.latitude >= 44 && record.latitude <= 53 && record.longitude >= 20 && record.longitude <= 42
      ).map((record) => ({ record, point: map.latLngToContainerPoint([record.latitude, record.longitude]) }))
        .filter(({ point }) => point.x >= -70 && point.y >= -70 && point.x <= size.x + 70 && point.y <= size.y + 70);

      const groups: number[][] = [];
      if (zoom < 15) {
        const radius = zoom < 9 ? 58 : zoom < 12 ? 48 : 40;
        const parents = visible.map((_item, index) => index);
        const find = (value: number): number => parents[value] === value ? value : (parents[value] = find(parents[value]));
        const join = (a: number, b: number) => { const ra = find(a); const rb = find(b); if (ra !== rb) parents[rb] = ra; };
        const buckets = new Map<string, number[]>();
        visible.forEach(({ point }, index) => {
          const cx = Math.floor(point.x / radius); const cy = Math.floor(point.y / radius);
          for (let dx = -1; dx <= 1; dx += 1) for (let dy = -1; dy <= 1; dy += 1) {
            for (const other of buckets.get(`${cx + dx}:${cy + dy}`) ?? []) {
              const candidate = visible[other].point;
              if (point.distanceTo(candidate) <= radius) join(index, other);
            }
          }
          const key = `${cx}:${cy}`;
          buckets.set(key, [...(buckets.get(key) ?? []), index]);
        });
        const grouped = new Map<number, number[]>();
        visible.forEach((_item, index) => {
          const root = find(index);
          grouped.set(root, [...(grouped.get(root) ?? []), index]);
        });
        groups.push(...grouped.values());
      } else {
        visible.forEach((_item, index) => groups.push([index]));
      }

      const duplicatePositions = new Map<string, number[]>();
      if (zoom >= 15) visible.forEach(({ record }, index) => {
        const key = `${record.latitude.toFixed(6)}:${record.longitude.toFixed(6)}`;
        duplicatePositions.set(key, [...(duplicatePositions.get(key) ?? []), index]);
      });
      const displayPositions = new Map<number, L.LatLng>();
      duplicatePositions.forEach((indices) => {
        if (indices.length < 2) return;
        const base = visible[indices[0]];
        const center = map.project([base.record.latitude, base.record.longitude], zoom);
        const radius = 14 + Math.min(14, indices.length * 1.2);
        indices.forEach((index, order) => {
          const angle = (Math.PI * 2 * order) / indices.length;
          const offset = L.point(center.x + Math.cos(angle) * radius, center.y + Math.sin(angle) * radius);
          displayPositions.set(index, map.unproject(offset, zoom));
        });
      });

      groups.forEach((indices) => {
        if (indices.length > 1 && zoom < 15) {
          const center = indices.reduce((acc, index) => acc.add(visible[index].point), L.point(0, 0)).divideBy(indices.length);
          const latlng = map.containerPointToLatLng(center);
          const clusterSize = indices.length >= 100 ? 40 : indices.length >= 10 ? 36 : 32;
          const icon = L.divIcon({
            className: "facility-cluster-host",
            html: `<span class="facility-cluster" style="--cluster-size:${clusterSize}px"><span>${indices.length}</span></span>`,
            iconSize: [clusterSize, clusterSize],
            iconAnchor: [clusterSize / 2, clusterSize / 2]
          });
          const marker = L.marker(latlng, { icon, title: `${indices.length} медичних закладів у групі`, keyboard: true });
          const bounds = L.latLngBounds(indices.map((index) => [visible[index].record.latitude, visible[index].record.longitude] as [number, number]));
          marker.on("click", () => map.fitBounds(bounds, { padding: [34, 34], maxZoom: 16 }));
          marker.addTo(layer);
          return;
        }

        const index = indices[0];
        const record = visible[index].record;
        const symbol = facilitySymbol([record.values.facility_type, record.values.name].filter(Boolean).join(" "));
        const name = String(record.values.name ?? "Медичний заклад");
        const html = `<span class="facility-marker facility-marker--${symbol.kind}" style="--marker-size:${iconSize}px"><span>${symbol.glyph}</span></span>`;
        const icon = L.divIcon({
          className: "facility-marker-host",
          html,
          iconSize: [iconSize, iconSize],
          iconAnchor: [iconSize / 2, iconSize / 2]
        });
        const marker = L.marker(displayPositions.get(index) ?? [record.latitude, record.longitude], { icon, title: name, keyboard: true });
        marker.bindTooltip(name, { direction: "top", offset: [0, -iconSize / 2] });
        const popup = document.createElement("div");
        const heading = document.createElement("strong");
        heading.textContent = name;
        popup.appendChild(heading);
        const details: Array<[string, unknown]> = [
          [translate("Тип"), record.values.facility_type || translate(symbol.label)],
          [translate("Адреса"), record.values.address],
          [translate("Населений пункт"), record.values.settlement],
          [translate("Громада"), record.values.hromada],
          [translate("Форма власності"), record.values.ownership],
          [translate("ЄДРПОУ"), record.values.edrpou],
          [translate("Телефон"), record.values.phone],
          [translate("Електронна пошта"), record.values.email],
          [translate("ID місця надання послуг НСЗУ"), record.values.division_id],
          [translate("Джерело"), record.values.source],
          [translate("Дата оновлення джерела"), record.values.source_updated_at]
        ];
        details.forEach(([label, value]) => {
          if (value === undefined || value === null || String(value).trim() === "") return;
          const line = document.createElement("div");
          line.textContent = `${label}: ${String(value)}`;
          popup.appendChild(line);
        });
        marker.bindPopup(popup, { maxWidth: 280 });
        marker.addTo(layer);
      });
    };
    render();
    map.on("zoomend moveend", render);
    return () => { map.off("zoomend moveend", render); map.removeLayer(layer); };
  }, [map, records]);
  return null;
}

function normalizeHeaderKey(value: unknown): string {
  return String(value ?? "")
    .trim()
    .toLocaleLowerCase("uk-UA")
    .normalize("NFKC")
    .replace(/[^\p{L}\p{N}]+/gu, "");
}

function parseSpreadsheetNumber(value: unknown): number {
  if (typeof value === "number") return value;
  const text = String(value ?? "").trim().replace(/[\s\u00a0]/g, "").replace(",", ".");
  if (!text) return Number.NaN;
  return Number(text);
}

function normalizeDate(value: unknown): string {
  if (value instanceof Date && Number.isFinite(value.getTime())) {
    return [value.getFullYear(), String(value.getMonth() + 1).padStart(2, "0"), String(value.getDate()).padStart(2, "0")].join("-");
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return [parsed.y, String(parsed.m).padStart(2, "0"), String(parsed.d).padStart(2, "0")].join("-");
  }
  const text = String(value ?? "").trim();
  const european = text.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})/);
  if (european) return `${european[3]}-${european[2].padStart(2, "0")}-${european[1].padStart(2, "0")}`;
  const iso = text.match(/^(\d{4}-\d{2}-\d{2})/);
  if (iso) return iso[1];
  if (/^\d+(?:[.,]\d+)?$/.test(text)) {
    const serial = Number(text.replace(",", "."));
    const parsed = XLSX.SSF.parse_date_code(serial);
    if (parsed) return [parsed.y, String(parsed.m).padStart(2, "0"), String(parsed.d).padStart(2, "0")].join("-");
  }
  return "";
}

function toCase(row: Record<string, unknown>, index: number): CaseRecord | null {
  const normalizedEntries = Object.entries(row).map(([key, value]) => [normalizeHeaderKey(key), value] as const);
  const get = (...keys: string[]) => {
    const aliases = new Set(keys.map(normalizeHeaderKey));
    const found = normalizedEntries.find(([key, value]) => Array.from(aliases).some((alias) => key === alias || (alias.length >= 4 && key.startsWith(alias))) && value !== undefined && value !== null && String(value).trim() !== "");
    return found?.[1];
  };
  const rawLatitude = get("latitude", "lat", "широта", "координата широти", "широта gps");
  const rawLongitude = get("longitude", "lon", "lng", "довгота", "координата довготи", "довгота gps");
  const latitude = parseSpreadsheetNumber(rawLatitude);
  const longitude = parseSpreadsheetNumber(rawLongitude);
  const disease = String(get("disease", "нозологія", "діагноз", "final_diagnosis", "інфекційне захворювання", "назва нозології") ?? "").trim();
  const date = normalizeDate(get("date", "дата", "дата захворювання", "onset_date", "дата початку", "дата реєстрації"));
  const hromada = String(get("hromada", "громада", "територія", "територіальна громада", "назва громади", "отг") ?? "").trim() || "Не визначено";
  // Keep epidemiological records even when location coordinates are absent or incomplete.
  if (!disease || !date) return null;
  const hasCoordinates = Number.isFinite(latitude) && Number.isFinite(longitude) &&
    latitude >= 44 && latitude <= 53 && longitude >= 20 && longitude <= 42;
  const rawAge = get("age", "вік");
  const parsedAge = rawAge === undefined ? Number.NaN : parseSpreadsheetNumber(rawAge);
  return {
    id: String(get("id", "номер", "case_id", "ідентифікатор") ?? ("IMP-" + String(index + 1).padStart(4, "0"))),
    disease, date, hromada,
    settlement: String(get("settlement", "населений пункт", "нп", "населений пункт / місто") ?? ""),
    latitude: hasCoordinates ? latitude : undefined,
    longitude: hasCoordinates ? longitude : undefined,
    status: String(get("status", "статус") ?? "Не визначено"),
    age: Number.isFinite(parsedAge) ? parsedAge : undefined,
    sex: String(get("sex", "стать") ?? ""),
    notes: String(get("notes", "примітки", "коментар") ?? "")
  };
}
function hasValidGeoJSONCoordinates(value: unknown): boolean {
  if (!Array.isArray(value) || value.length === 0) return false;
  if (typeof value[0] === "number" && typeof value[1] === "number") {
    const [longitude, latitude] = value;
    return Number.isFinite(longitude) && Number.isFinite(latitude) &&
      longitude >= -180 && longitude <= 180 && latitude >= -90 && latitude <= 90;
  }
  return value.every(hasValidGeoJSONCoordinates);
}


const UI_TRANSLATIONS: Record<string, string> = {
  "ЕПІДЕМІОЛОГІЧНЕ КАРТОГРАФУВАННЯ": "EPIDEMIOLOGICAL MAPPING",
  "Локальний режим": "Local mode",
  "Імпорт даних": "Import data",
  "ТЕРНОПІЛЬСЬКА ОБЛАСТЬ / ЕПІДНАГЛЯД": "TERNOPIL REGION / SURVEILLANCE",
  "Карта епідеміологічної ситуації": "Epidemiological situation map",
  "Просторовий аналіз випадків, часові фільтри та огляд територій.": "Spatial case analysis, time filters and geographic overview.",
  "Експортувати вибірку": "Export selection",
  "Закрити повідомлення": "Close message",
  "Імпорт книги EpiGIS або CSV НСЗУ": "Import EpiGIS workbook or NHSU CSV",
  "Завантажити шаблон / сховище": "Download template / data store",
  "Випадків у вибірці": "Cases in selection",
  "За активними фільтрами": "Using active filters",
  "Підтверджено": "Confirmed",
  "Немає записів": "No records",
  "Громад із випадками": "Communities with cases",
  "У межах поточного відбору": "Within current selection",
  "Нозологій": "Diagnoses",
  "Різних діагнозів у вибірці": "Distinct diagnoses in selection",
  "Фільтри": "Filters",
  "Скинути": "Reset",
  "Пошук за записом або НП": "Search records or settlement",
  "Діагноз, громада, населений пункт…": "Diagnosis, community, settlement…",
  "Нозології (можна обрати декілька)": "Diagnoses (multiple selection)",
  "Усі нозології": "All diagnoses",
  "Очистити вибір": "Clear selection",
  "Територіальні громади (можна обрати декілька)": "Territorial communities (multiple selection)",
  "Усі громади": "All communities",
  "Від дати": "From date",
  "До дати": "To date",
  "Лише підтверджені": "Confirmed only",
  "Фільтри предметних реєстрів": "Thematic registry filters",
  "Тип реєстру (можна обрати декілька)": "Registry type (multiple selection)",
  "Реєстри не вибрано": "No registries selected",
  "За замовчуванням: реєстр випадків": "Default: case registry",
  "Реєстр випадків": "Case registry",
  "Пошук об’єкта / атрибута": "Search object / attribute",
  "Назва, адреса, оператор, вид…": "Name, address, operator, type…",
  "Територіальна громада": "Territorial community",
  "Усі типи": "All types",
  "Знайдено об’єктів реєстрів: ": "Registry objects found: ",
  "Умовні позначення": "Legend",
  "Підтверджений випадок": "Confirmed case",
  "Розслідується": "Under investigation",
  "Предметні реєстри": "Thematic registries",
  "Тип реєстру для наступного GeoJSON-імпорту": "Registry for the next GeoJSON import",
  "Реєстр: ": "Registry: ",
  "записів · окремий набір даних": "records · separate dataset",
  "атрибутів · ": "attributes · ",
  "різні геометрії": "mixed geometries",
  "Реєстр випадків доступний у таблиці «Реєстр випадків» нижче та через аркуш Cases в Excel-сховищі. Записи випадків не дублюються в предметних реєстрах.": "The case registry is available in the Case Registry table below and in the Cases sheet of the Excel data store. Case records are not duplicated in thematic registries.",
  "Додати запис": "Add record",
  "Редагування запису": "Edit record",
  "Новий запис": "New record",
  "Вибрати…": "Select…",
  "Широта": "Latitude",
  "Довгота": "Longitude",
  "Зберегти": "Save",
  "Скасувати": "Cancel",
  "Редагувати запис": "Edit record",
  "Видалити запис": "Delete record",
  "Записи зберігаються локально в цьому браузері. Координати використовуються для відображення точок на карті. Для ліній і полігонів імпортуй GeoJSON.": "Records are stored locally in this browser. Coordinates are used to display map points. Import GeoJSON for lines and polygons.",
  "Географічний огляд": "Geographic overview",
  "точкових записів на карті": "point records on map",
  "За вибраними параметрами записів не знайдено.": "No records found for the selected criteria.",
  "Координати демонстраційних записів умовні.": "Demo record coordinates are illustrative.",
  "Лікарні": "Hospitals",
  "Первинна допомога": "Primary care",
  "Стоматології": "Dental clinics",
  "Екстрена допомога": "Emergency care",
  "Лабораторії": "Laboratories",
  "Приватні центри": "Private clinics",
  "Результати просторового аналізу": "Spatial analysis results",
  "Закрити таблицю": "Close table",
  "Тип геометрії": "Geometry type",
  "Кількість перетинів": "Intersection count",
  "Атрибути": "Attributes",
  "Немає атрибутів": "No attributes",
  "Таблиця відображає геометричні перетини, а не підтверджений епідеміологічний зв’язок. Для повного набору атрибутів експортуй GeoJSON відповідного шару.": "This table shows geometric intersections, not a confirmed epidemiological relationship. Export the relevant GeoJSON layer for the full attribute set.",
  "Натисни на рядок, щоб переглянути деталі": "Select a row to view details",
  "Реєстр": "Registry",
  "Об’єкт": "Object",
  "Громада / НП": "Community / settlement",
  "Тип / характеристика": "Type / characteristic",
  "Дата": "Date",
  "Нозологія": "Diagnosis",
  "Статус": "Status",
  "Показано 100 із ": "Showing 100 of ",
  "записів. Експорт містить усю вибірку.": "records. Export includes the full selection.",
  "Немає записів для відображення.": "No records to display.",
  "Об’єкти предметних реєстрів": "Thematic registry objects",
  "За активними фільтрами реєстрів": "Using active registry filters",
  "об’єктів. Повні дані містяться в Excel-сховищі.": "objects. Full data is available in the Excel data store.",
  "Об’єктів для відображення немає.": "No objects to display.",
  "Керування картою": "Map controls",
  "Географічні шари": "Geographic layers",
  "Додати GeoJSON": "Add GeoJSON",
  "Поки немає власних шарів. Додай GeoJSON із точками, лініями або полігонами.": "No custom layers yet. Add GeoJSON containing points, lines or polygons.",
  "Показати шар на карті": "Show layer on map",
  "Перейменувати шар": "Rename layer",
  "Експортувати GeoJSON": "Export GeoJSON",
  "Видалити шар": "Delete layer",
  "Креслення на карті": "Map drawing",
  "Точка": "Point",
  "Лінія": "Line",
  "Полігон": "Polygon",
  "Клацни на карті, щоб поставити точку.": "Click the map to place a point.",
  "Познач вершини на карті. Додано: ": "Mark vertices on the map. Added: ",
  "Завершити": "Finish",
  "Буферний аналіз": "Buffer analysis",
  "Вихідний шар": "Source layer",
  "Вибрати шар…": "Select layer…",
  "Радіус буфера, метрів": "Buffer radius, metres",
  "Побудувати буфер": "Create buffer",
  "Просторові перетини": "Spatial intersections",
  "Зона аналізу (наприклад, буфер)": "Analysis zone (e.g. buffer)",
  "Вибрати шар зоною…": "Select zone layer…",
  "Шар об'єктів для перевірки": "Layer of objects to check",
  "Вибрати шар об'єктів…": "Select object layer…",
  "Знайти об'єкти в зоні": "Find objects in zone",
  "Перевіряється геометричний перетин об'єктів двох шарів. Результат додається окремим шаром; це не є автоматичним епідеміологічним висновком.": "Geometric intersections between two layers are checked. The result is added as a separate layer; this is not an automatic epidemiological conclusion.",
  "Шари зберігаються локально у браузері. Вони не передаються на сервер і можуть бути недоступні на іншому пристрої.": "Layers are stored locally in this browser. They are not sent to a server and may not be available on another device.",
  "Картка випадку": "Case details",
  "Дата реєстрації": "Registration date",
  "Територія": "Territory",
  "Населений пункт": "Settlement",
  "Вік": "Age",
  "Стать": "Sex",
  "Не вказано": "Not specified",
  "років": "years",
  "Примітки": "Notes",
  "Випадок не вибрано": "No case selected",
  "Обери точку на карті або запис у реєстрі, щоб побачити деталі.": "Select a point on the map or a registry record to view details.",
  "Важливо": "Important",
  "Просторове скупчення точок саме по собі не підтверджує спалах. Потрібна епідеміологічна оцінка.": "A spatial cluster of points alone does not confirm an outbreak. Epidemiological assessment is required.",
  "Демонстраційні дані · Не для клінічних рішень": "Demonstration data · Not for clinical decisions",
  "Розробник:": "Developer:",
  "Можна завантажити повну Excel-книгу EpiGIS або безпосередньо pmg_legal_entity_divisions_info.csv. Для CSV система відбере активні місця надання послуг Тернопільської області, зіставить поля та імпортує записи з валідними координатами.": "You can upload the complete EpiGIS Excel workbook or the pmg_legal_entity_divisions_info.csv file directly. For CSV, the system selects active service locations in Ternopil Region, maps the fields and imports records with valid coordinates.",
  "+ Додати запис": "+ Add record",
  "Широта *": "Latitude *",
  "Довгота *": "Longitude *",
  "Вулична карта": "Street map",
  "Супутник": "Satellite",
  "Карта онлайн": "Online map",
  "Заклади охорони здоров'я": "Healthcare facilities",
  "Лікарні, амбулаторії, лабораторії, пункти екстреної допомоги.": "Hospitals, outpatient clinics, laboratories and emergency care stations.",
  "Водна інфраструктура": "Water infrastructure",
  "Водозабори, водогони, каналізаційні мережі та очисні споруди.": "Water intakes, water pipelines, sewer networks and wastewater treatment plants.",
  "Харчові об'єкти": "Food facilities",
  "Підприємства харчування, виробництва, торгівлі та водопостачання харчових об'єктів.": "Food service, production, retail and food-facility water supply sites.",
  "Природні об'єкти": "Natural features",
  "Річки, водойми, ліси, заболочені території та інші природні об'єкти.": "Rivers, water bodies, forests, wetlands and other natural features.",
  "Епідеміологічні вогнища": "Epidemiological foci",
  "Вогнища, території розслідування та групові випадки; персональні дані не передбачені.": "Foci, investigation areas and case clusters; personal data are not included.",
  "Біологічний моніторинг": "Biological monitoring",
  "Місця відбору проб, кліщі, гризуни, переносники та результати лабораторних досліджень.": "Sampling sites, ticks, rodents, vectors and laboratory test results.",
  "Метеорологія": "Meteorology",
  "Метеостанції та спостереження: температура, опади, вологість, вітер.": "Weather stations and observations: temperature, precipitation, humidity and wind.",
  "Населення": "Population",
  "Агрегована чисельність населення за населеними пунктами та громадами, без персональних записів.": "Aggregated population counts by settlement and community, without personal records.",
  "Ідентифікатор": "Identifier",
  "Назва об'єкта": "Object name",
  "Адреса": "Address",
  "Джерело даних": "Data source",
  "Дата актуалізації": "Last updated",
  "Тип закладу": "Facility type",
  "Обласна лікарня": "Regional hospital",
  "Міська / районна лікарня": "City / district hospital",
  "Багатопрофільна лікарня": "Multidisciplinary hospital",
  "Спеціалізована лікарня": "Specialized hospital",
  "Дитяча лікарня": "Children's hospital",
  "Пологовий будинок": "Maternity hospital",
  "ЦПМСД": "Primary healthcare centre",
  "Амбулаторія загальної практики": "General practice outpatient clinic",
  "ФАП / ФП": "Feldsher-midwife station / feldsher station",
  "Стоматологія": "Dental clinic",
  "Приватний медичний центр": "Private medical centre",
  "Діагностичний центр": "Diagnostic centre",
  "Лабораторія": "Laboratory",
  "Реабілітаційний центр": "Rehabilitation centre",
  "Центр екстреної медичної допомоги": "Emergency medical care centre",
  "Пункт базування ЕМД": "Emergency medical team base",
  "Інше": "Other",
  "Форма власності": "Ownership type",
  "ЄДРПОУ": "EDRPOU code",
  "ID юридичної особи НСЗУ": "NHSU legal entity ID",
  "ID місця надання послуг": "Service location ID",
  "Тип підрозділу НСЗУ": "NHSU division type",
  "Телефон": "Phone",
  "Електронна пошта": "Email",
  "Дата оновлення джерела": "Source update date",
  "Статус ліцензії МОЗ": "Ministry of Health licence status",
  "Тип об'єкта": "Object type",
  "Водозабір": "Water intake",
  "Водогін": "Water pipeline",
  "Каналізація": "Sewerage",
  "Очисні споруди": "Wastewater treatment plant",
  "Балансоутримувач / оператор": "Asset owner / operator",
  "Вид діяльності": "Activity type",
  "Категорія ризику": "Risk category",
  "Тип природного об'єкта": "Natural feature type",
  "Дата виявлення": "Detection date",
  "Активне": "Active",
  "Під наглядом": "Under surveillance",
  "Закрите": "Closed",
  "Дата відбору": "Sampling date",
  "Вид / таксон": "Species / taxon",
  "Стадія розвитку": "Life stage",
  "Результат дослідження": "Test result",
  "Дата і час спостереження": "Observation date and time",
  "Температура, °C": "Temperature, °C",
  "Опади, мм": "Precipitation, mm",
  "Вологість, %": "Humidity, %",
  "Швидкість вітру, м/с": "Wind speed, m/s",
  "Чисельність населення": "Population size",
  "Рік оцінки": "Reference year",
  "Тип": "Type",
  "Медичний заклад": "Healthcare facility",
  "Інший медичний заклад": "Other healthcare facility",
  "Лікарня": "Hospital",
  "Медичний центр / клініка": "Medical centre / clinic",
  "ID місця надання послуг НСЗУ": "NHSU service location ID",
  "Аркуш випадків знайдено, але жоден рядок не пройшов перевірку обов’язкових полів disease і date. Перевір значення у рядках; заголовки: ": "The Cases sheet was found, but no rows passed validation of the required disease and date fields. Check the row values; headers: ",
  "не визначені": "not identified",
  ". Попередні випадки збережено.": ". Previous cases were preserved.",
  " Випадки не змінювалися.": " Cases were not changed.",
  " Предметні реєстри не змінювалися.": " Thematic registries were not changed.",
  " Геошари не змінювалися.": " Geographic layers were not changed.",
  " Геошари не оновлено через помилки GeoJSON; попередні шари збережено.": " Geographic layers were not updated due to GeoJSON errors; previous layers were preserved.",
  " Дані збережено в локальному сховищі цього браузера.": " Data have been saved to this browser's local storage.",
  "широта": "latitude",
  "довгота": "longitude",
  "громада": "community",
  "територіальна громада": "territorial community",
  "населений пункт": "settlement",
  "дата реєстрації": "registration date",
  "дата захворювання": "onset date",
  "нозологія": "diagnosis",
  "діагноз": "diagnosis",
  "статус": "status",
  "вік": "age",
  "стать": "sex",
  "примітки": "notes",
  "Експорт": "Export",
  "Імпорт": "Import",
  "Видалити": "Delete",
  "Перейменувати": "Rename",
  "Завантажити": "Download",
  "Переглянути": "View",
  "Додати": "Add",
  "Так": "Yes",
  "Ні": "No",
  "Не вдалося зберегти випадки: сховище браузера переповнене або недоступне.": "Could not save cases: browser storage is full or unavailable.",
  "Не вдалося зберегти реєстр: сховище браузера переповнене або недоступне.": "Could not save the registry: browser storage is full or unavailable.",
  "Не вдалося зберегти географічні шари: сховище браузера переповнене або недоступне.": "Could not save geographic layers: browser storage is full or unavailable.",
  "Не вдалося завантажити межі громад. Перевір наявність файлу public/data/ternopil-hromadas.geojson у збірці.": "Could not load community boundaries. Check that public/data/ternopil-hromadas.geojson is included in the build.",
  "У книзі не знайдено аркушів, сумісних з імпортом EpiGIS.": "No worksheets compatible with EpiGIS import were found in the workbook.",
  "Не вдалося прочитати книгу. Перевір формат і заголовки; підтримуються книги EpiGIS, CSV НСЗУ з місцями надання послуг та таблиці випадків.": "Could not read the workbook. Check the format and headers. Supported formats: EpiGIS workbooks, NHSU service-location CSV files and case tables.",
  "Файл має містити GeoJSON Feature або FeatureCollection з геометріями.": "The file must contain a GeoJSON Feature or FeatureCollection with geometries.",
  "У файлі не знайдено підтримуваних геометрій: точки, лінії або полігони.": "No supported geometries were found in the file: points, lines or polygons.",
  "Не вдалося прочитати GeoJSON. Перевір структуру та синтаксис JSON.": "Could not read GeoJSON. Check the JSON structure and syntax.",
  "Потрібні назва об\u0027єкта та коректні координати.": "An object name and valid coordinates are required.",
  "Заповни всі обов\u0027язкові поля реєстру.": "Complete all required registry fields.",
  "Запис реєстру оновлено.": "Registry record updated.",
  "Новий запис додано до реєстру.": "New record added to the registry.",
  "Видалити цей запис із локального реєстру?": "Delete this record from the local registry?",
  "Нова назва географічного шару:": "New geographic layer name:",
  "Для лінії потрібно щонайменше 2 вершини, для полігона — 3.": "A line requires at least 2 vertices; a polygon requires at least 3.",
  "Спочатку вибери шар для буферизації.": "First select a layer to buffer.",
  "Вкажи відстань від 1 до 100 000 метрів.": "Enter a distance from 1 to 100,000 metres.",
  "Не вдалося побудувати буфер для вибраного шару.": "Could not create a buffer for the selected layer.",
  "Помилка буферизації. Перевір геометрії у вихідному GeoJSON.": "Buffering error. Check the geometries in the source GeoJSON.",
  "Вибери шар зоною аналізу та шар об\u0027єктів для перевірки.": "Select an analysis-zone layer and an object layer to check.",
  "Для аналізу потрібно вибрати два різні шари.": "Two different layers must be selected for the analysis.",
  "Один із вибраних шарів не містить об\u0027єктів.": "One of the selected layers contains no objects.",
  "Не вдалося виконати перетин. Перевір валідність геометрій у двох GeoJSON-шарах.": "Could not perform the intersection. Check the geometry validity in both GeoJSON layers.",
  "Не вдалося створити зображення карти. Перевір доступність картографічних тайлів і повтори спробу.": "Could not create the map image. Check map tile availability and try again.",
  "Підкладка карти": "Map basemap",
  "Формат експорту карти": "Map export format",
  "Створення…": "Creating…",
  "Зберегти карту": "Save map",
  "Демонстраційні дані — не реальні випадки.": "Demonstration data — not real cases.",
  "атрибутів · різні геометрії": "attributes · mixed geometries",
  "PNG": "PNG",
  "JPEG": "JPEG",
  "PDF": "PDF",
  "об’єктів": "objects",
  "ID": "ID",
  " записів. Експорт містить усю вибірку.": " records. Export includes the full selection.",
  " об’єктів. Повні дані містяться в Excel-сховищі.": " objects. Full data is available in the Excel data store.",
  ["Мова"]: "Language",
  ["Обрано: "]: "Selected: ",
  [" район"]: " district",
  ["записів"]: "records",
  ["об'єктів"]: "objects",
  ["об’єктів шару"]: "layer objects",
  ["об’єктів перевірено щодо шару"]: "objects checked against layer",
  ["Перетинів не виявлено: "]: "No intersections found: ",
  ["Побудовано буфер "]: "Buffer created: ",
  [" м для "]: " m for ",
  [" об'єктів шару "]: " objects in layer ",
  ["Створено шар "]: "Layer created: ",
  ["Шар перейменовано на "]: "Layer renamed to ",
  ["Додано шар "]: "Layer added: ",
  ["Відхилено некоректних або непідтримуваних: "]: "Invalid or unsupported features rejected: ",
  ["Громаду не вказано"]: "Community not specified",
  ["Об'єкт без назви"]: "Unnamed object",
  ["Об’єкт реєстру"]: "Registry object",
  ["Об'єкт шару"]: "Layer object",
  ["Вибери шар для буферизації."]: "Select a layer to buffer.",
  ["Координати"]: "Coordinates",
  ["Демонстраційні дані"]: "Demonstration data",
  ["Масштабувати карту до шару"]: "Zoom to layer",
  "Можна завантажити повну Excel-книгу EpiGIS або безпосередньо pmg_legal_entity_divisions_info.csv.": "Upload a complete EpiGIS Excel workbook or the pmg_legal_entity_divisions_info.csv file directly.",
  "доступний у таблиці «Реєстр випадків» нижче та через аркуш Cases в Excel-сховищі. Записи випадків не дублюються в предметних реєстрах.": "is available in the Case Registry table below and in the Cases sheet of the Excel data store. Case records are not duplicated in thematic registries.",
  "Імпорт завершено.": "Import completed.",
  " Випадки: ": " Cases: ",
  " (без координат: ": " (without coordinates: ",
  ", відхилено: ": ", rejected: ",
  " Предметні реєстри: ": " Thematic registries: ",
  " записів у ": " records across ",
  " аркушах.": " worksheets.",
  " Геошари: ": " Geographic layers: ",
  " шарів.": " layers.",
  " Некоректні рядки реєстрів/геошарів: ": " Invalid registry/geographic-layer rows: ",
  "Потрібні назва об'єкта та коректні координати.": "An object name and valid coordinates are required.",
  "Шар перейменовано на «": "Layer renamed to “",
  "Створено шар «": "Created layer “",
  "Вибери шар зоною аналізу та шар об'єктів для перевірки.": "Select an analysis-zone layer and a target layer to check.",
  "Один із вибраних шарів не містить об'єктів.": "One of the selected layers contains no features.",
  "Не вдалося завантажити випадки: сховище браузера переповнене або недоступне.": "Could not save cases: browser storage is full or unavailable.",
};

function localizeNode(node: React.ReactNode, language: "uk" | "en"): React.ReactNode {
  if (language === "uk") return node;
  if (typeof node === "string") return UI_TRANSLATIONS[node] ?? node;
  if (typeof node === "number" || node === null || node === undefined || typeof node === "boolean") return node;
  if (Array.isArray(node)) return node.map((child) => localizeNode(child, language));
  if (isValidElement(node)) {
    const props: Record<string, unknown> = { ...node.props };
    for (const key of ["placeholder", "title", "aria-label", "label", "description"]) {
      if (typeof props[key] === "string") props[key] = UI_TRANSLATIONS[props[key] as string] ?? props[key];
    }
    // Translate only explicit UI option arrays; never alter arbitrary record data.
    if (Array.isArray(props.options)) {
      props.options = (props.options as unknown[]).map((option) =>
        typeof option === "string" ? (UI_TRANSLATIONS[option] ?? option) : option
      );
    }
    if (props.children !== undefined) props.children = localizeNode(props.children as React.ReactNode, language);
    return cloneElement(node, props);
  }
  return node;
}

export default function App() {
  const [language, setLanguage] = useState<"uk" | "en">(() => localStorage.getItem("ternopil-epigis-language") === "en" ? "en" : "uk");
  useEffect(() => { localStorage.setItem("ternopil-epigis-language", language); }, [language]);
  const [cases, setCases] = useState<CaseRecord[]>(() => { try { const saved = localStorage.getItem("ternopil-epigis-cases-v1"); return saved ? JSON.parse(saved) as CaseRecord[] : DEMO_CASES; } catch { return DEMO_CASES; } });
  const [gisLayers, setGisLayers] = useState<GISLayer[]>(() => {
    try { const saved = localStorage.getItem("ternopil-epigis-layers-v1"); return saved ? JSON.parse(saved) as GISLayer[] : []; } catch { return []; }
  });
  const [layerMessage, setLayerMessage] = useState("");
  const [mapExportFormat, setMapExportFormat] = useState<"png" | "jpeg" | "pdf">("png");
  const [mapExporting, setMapExporting] = useState(false);
  const [mapBaseLayer, setMapBaseLayer] = useState<"street" | "satellite">("street");
  const mapExportRef = useRef<HTMLDivElement>(null);
  const [activeRegistryId, setActiveRegistryId] = useState("healthcare");
  const [registryRecords, setRegistryRecords] = useState<RegistryRecord[]>(() => { try { const saved = localStorage.getItem("ternopil-epigis-registry-records-v1"); return saved ? JSON.parse(saved) as RegistryRecord[] : []; } catch { return []; } });
  const [registryFormOpen, setRegistryFormOpen] = useState(false);
  const [editingRegistryRecordId, setEditingRegistryRecordId] = useState("");
  const [registryFormValues, setRegistryFormValues] = useState<Record<string, string>>({});
  const [fitLayerId, setFitLayerId] = useState("");
  const [intersectionResultLayerId, setIntersectionResultLayerId] = useState("");
  const [drawMode, setDrawMode] = useState<DrawMode>("none");
  const [drawPoints, setDrawPoints] = useState<[number, number][]>([]);
  const [bufferLayerId, setBufferLayerId] = useState("");
  const [bufferMeters, setBufferMeters] = useState("500");
  const [intersectionZoneId, setIntersectionZoneId] = useState("");
  const [intersectionTargetId, setIntersectionTargetId] = useState("");
  const layerFileRef = useRef<HTMLInputElement>(null);
  const [hromadaBoundaries, setHromadaBoundaries] = useState<any>(null);
  const [selectedDiseases, setSelectedDiseases] = useState<string[]>([]);
  const [selectedHromadas, setSelectedHromadas] = useState<string[]>([]);
  const [fromDate, setFromDate] = useState("2026-01-01");
  const [toDate, setToDate] = useState("2026-12-31");
  const [search, setSearch] = useState("");
  const [selected, setSelected] = useState<CaseRecord | null>(null);
  const [message, setMessage] = useState("Демонстраційні дані — не реальні випадки.");
  const [showUpload, setShowUpload] = useState(false);
  const [onlyConfirmed, setOnlyConfirmed] = useState(false);
  const [selectedRegistryIds, setSelectedRegistryIds] = useState<string[]>(["cases"]);
  const [registryCommunityFilter, setRegistryCommunityFilter] = useState("Усі громади");
  const [registryTypeFilter, setRegistryTypeFilter] = useState("Усі типи");
  const [registrySearch, setRegistrySearch] = useState("");
  const [registryFromDate, setRegistryFromDate] = useState("");
  const [registryToDate, setRegistryToDate] = useState("");
  const fileRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    try { localStorage.setItem("ternopil-epigis-cases-v1", JSON.stringify(cases)); } catch { setLayerMessage("Не вдалося зберегти випадки: сховище браузера переповнене або недоступне."); }
  }, [cases]);

  useEffect(() => {
    try { localStorage.setItem("ternopil-epigis-registry-records-v1", JSON.stringify(registryRecords)); } catch { setLayerMessage("Не вдалося зберегти реєстр: сховище браузера переповнене або недоступне."); }
  }, [registryRecords]);

  useEffect(() => {
    try {
      localStorage.setItem("ternopil-epigis-layers-v1", JSON.stringify(gisLayers));
    } catch {
      setLayerMessage("Не вдалося зберегти географічні шари: сховище браузера переповнене або недоступне.");
    }
  }, [gisLayers]);

  useEffect(() => {
    let cancelled = false;
    fetch("/data/ternopil-hromadas.geojson")
      .then((response) => {
        if (!response.ok) throw new Error("GeoJSON unavailable");
        return response.json();
      })
      .then((data) => {
        if (!cancelled && data?.type === "FeatureCollection" && Array.isArray(data.features)) {
          setHromadaBoundaries(data);
        }
      })
      .catch(() => {
        if (!cancelled) setMessage("Не вдалося завантажити межі громад. Перевір наявність файлу public/data/ternopil-hromadas.geojson у збірці.");
      });
    return () => { cancelled = true; };
  }, []);

  const diseases = useMemo(() => Array.from(new Set(cases.map((item) => item.disease))).sort(), [cases]);
  const hromadas = useMemo(() => Array.from(new Set(cases.map((item) => item.hromada))).sort(), [cases]);
  const filtered = useMemo(() => cases.filter((item) =>
    selectedRegistryIds.includes("cases") &&
    (selectedDiseases.length === 0 || selectedDiseases.includes(item.disease)) &&
    (selectedHromadas.length === 0 || selectedHromadas.includes(item.hromada)) &&
    (!fromDate || item.date >= fromDate) && (!toDate || item.date <= toDate) &&
    (!onlyConfirmed || item.status === "Підтверджено") &&
    (!search || [item.id, item.settlement, item.hromada, item.disease].join(" ").toLowerCase().includes(search.toLowerCase())) &&
    (!selectedRegistryIds.includes("cases") || registryCommunityFilter === "Усі громади" || item.hromada === registryCommunityFilter) &&
    (!selectedRegistryIds.includes("cases") || registryTypeFilter === "Усі типи" || item.disease === registryTypeFilter) &&
    (!selectedRegistryIds.includes("cases") || !registryFromDate || item.date >= registryFromDate) &&
    (!selectedRegistryIds.includes("cases") || !registryToDate || item.date <= registryToDate) &&
    (!selectedRegistryIds.includes("cases") || !registrySearch || [item.id, item.settlement, item.hromada, item.disease, item.status].join(" ").toLocaleLowerCase("uk-UA").includes(registrySearch.toLocaleLowerCase("uk-UA")))
  ), [cases, selectedRegistryIds, registryCommunityFilter, registryTypeFilter, registryFromDate, registryToDate, registrySearch, selectedDiseases, selectedHromadas, fromDate, toDate, onlyConfirmed, search]);

  const confirmed = filtered.filter((item) => item.status === "Підтверджено").length;
  const locations = new Set(filtered.map((item) => item.hromada)).size;
  const activeDiseases = new Set(filtered.map((item) => item.disease)).size;

  const importFile = async (file?: File) => {
    if (!file) return;
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
      const findSheetName = (sheetName: string, aliases: string[] = []) => {
        const wanted = new Set([sheetName, ...aliases].map(normalizeHeaderKey));
        return workbook.SheetNames.find((candidate) => wanted.has(normalizeHeaderKey(candidate)));
      };
      const readRows = (sheetName: string, headerGroups: string[][] = [], sheetAliases: string[] = []): Record<string, unknown>[] | null => {
        const actualSheetName = findSheetName(sheetName, sheetAliases);
        if (!actualSheetName) return null;
        const sheet = workbook.Sheets[actualSheetName];
        const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", blankrows: true, raw: true });
        let headerRow = 0;
        if (headerGroups.length) {
          let bestScore = 0;
          const scanLimit = Math.min(matrix.length, 20);
          for (let rowIndex = 0; rowIndex < scanLimit; rowIndex += 1) {
            const cells = new Set((matrix[rowIndex] ?? []).map(normalizeHeaderKey).filter(Boolean));
            const score = headerGroups.filter((group) => group.some((alias) => Array.from(cells).some((cell) => cell === normalizeHeaderKey(alias) || (normalizeHeaderKey(alias).length >= 4 && cell.startsWith(normalizeHeaderKey(alias)))))).length;
            if (score > bestScore) {
              bestScore = score;
              headerRow = rowIndex;
            }
          }
        }
        const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { range: headerRow, defval: "", raw: true });
        return rows.filter((row) => Object.values(row).some((value) => String(value ?? "").trim() !== ""));
      };
      const caseHeaderGroups = [
        ["disease", "нозологія", "діагноз", "final_diagnosis", "інфекційне захворювання", "назва нозології"],
        ["date", "дата", "дата захворювання", "onset_date", "дата початку", "дата реєстрації"],
        ["hromada", "громада", "територія", "територіальна громада", "назва громади", "отг"],
        ["latitude", "lat", "широта", "широта gps"],
        ["longitude", "lon", "lng", "довгота", "довгота gps"]
      ];
      const hasCasesSheet = Boolean(findSheetName("Cases", ["Випадки", "Випадки інфекційної захворюваності"]));
      const nszuHeaderGroups = [
        ["division_id", "ідентифікатор підрозділу"],
        ["division_name", "назва підрозділу"],
        ["division_type", "тип підрозділу"],
        ["residence_region", "регіон місцезнаходження"],
        ["lat", "latitude", "широта"],
        ["lng", "longitude", "довгота"]
      ];
      const singleSheetRows = !hasCasesSheet && workbook.SheetNames.length === 1
        ? readRows(workbook.SheetNames[0], nszuHeaderGroups)
        : null;
      const isNszuDivisionsFile = Boolean(singleSheetRows?.length && Object.keys(singleSheetRows[0]).some((key) => ["division_id", "ідентифікатор підрозділу"].includes(normalizeHeaderKey(key))));
      const isLegacySingleSheet = !hasCasesSheet && workbook.SheetNames.length === 1 && !isNszuDivisionsFile;
      const caseRows = hasCasesSheet
        ? readRows("Cases", caseHeaderGroups, ["Випадки", "Випадки інфекційної захворюваності"]) ?? []
        : isLegacySingleSheet
          ? readRows(workbook.SheetNames[0], caseHeaderGroups)
          : null;
      let caseCount = 0;
      let invalidCases = 0;
      let casesUpdated = false;
      let noCoordinateCases = 0;
      let caseImportWarning = "";
      if (caseRows !== null) {
        const actualCaseRows = caseRows.filter((row) => {
          const id = String(row.id ?? row.ID ?? "").trim();
          const notes = String(row.notes ?? row.примітки ?? "").toLocaleLowerCase("uk-UA");
          return !(id === "CASE-0001" && (notes.includes("приклад") || notes.includes("зразок")));
        });
        const parsedCases = actualCaseRows.map(toCase).filter((item): item is CaseRecord => item !== null);
        if (parsedCases.length > 0) {
          setCases(parsedCases);
          casesUpdated = true;
          caseCount = parsedCases.length;
          invalidCases = actualCaseRows.length - parsedCases.length;
          noCoordinateCases = parsedCases.filter((item) => !Number.isFinite(item.latitude) || !Number.isFinite(item.longitude)).length;
        } else if (caseRows.length > 0) {
          const detectedHeaders = Object.keys(caseRows[0] ?? {}).join(", ");
          const onlyExampleRows = actualCaseRows.length === 0 && caseRows.every((row) => {
            const id = String(row.id ?? row.ID ?? "").trim().toLocaleUpperCase("uk-UA");
            const notes = String(row.notes ?? row.примітки ?? "").toLocaleLowerCase("uk-UA");
            return id === "CASE-0001" && (notes.includes("приклад") || notes.includes("зразок"));
          });
          caseImportWarning = onlyExampleRows
            ? ' Аркуш Cases містить лише демонстраційний рядок CASE-0001 ("Приклад — замінити даними"). Його пропущено; додай реальні записи починаючи з наступного рядка. Попередні випадки збережено.'
            : " Аркуш випадків знайдено, але жоден рядок не пройшов перевірку обов’язкових полів disease і date. Перевір значення у рядках; заголовки: " + (detectedHeaders || "не визначені") + ". Попередні випадки збережено.";
        }
      }

      const replacedRegistryIds: string[] = [];
      const importedRegistryRecords: RegistryRecord[] = [];
      let invalidRegistryRows = 0;
      REGISTRIES.forEach((registry) => {
        const sheetName = REGISTRY_SHEET_NAMES[registry.id];
        const registryHeaderGroups = [
          ...registry.fields.map((field) => [field.key, field.label]),
          ["latitude", "широта", "lat"],
          ["longitude", "довгота", "lon", "lng"],
          ["record_id", "recordId", "_record_id"]
        ];
        const rows = registry.id === "healthcare" && isNszuDivisionsFile
          ? singleSheetRows
          : readRows(sheetName, registryHeaderGroups, [registry.label]);
        if (rows === null || rows.length === 0) return;
        const stagedRecords: RegistryRecord[] = [];
        let sheetInvalidRows = 0;
        rows.forEach((row, index) => {
          const nszuAliases: Record<string, string[]> = {
            object_id: ["division_id", "Ідентифікатор підрозділу"],
            name: ["division_name", "Назва підрозділу"],
            hromada: ["residence_gromada_name", "Назва громади місцезнаходження"],
            settlement: ["residence_settlement", "Населений пункт місцезнаходження", "residence_city"],
            address: ["residence_addresses", "Адреси місцезнаходження"],
            source: ["НСЗУ — pmg_legal_entity_divisions_info.csv"],
            facility_type: ["division_type", "Тип підрозділу"],
            legal_entity_id: ["legal_entity_id", "Ідентифікатор юридичної особи"],
            division_id: ["division_id", "Ідентифікатор підрозділу"],
            service_location_type: ["division_type", "Тип підрозділу"],
            phone: ["division_phone", "Телефон підрозділу"],
            email: ["division_email", "Електронна пошта підрозділу"]
          };
          const getValue = (...keys: string[]) => {
            const expandedKeys = registry.id === "healthcare"
              ? keys.flatMap((key) => [key, ...(nszuAliases[key] ?? [])])
              : keys;
            for (const key of expandedKeys) {
              const normalizedKey = normalizeHeaderKey(key);
              const found = Object.keys(row).find((candidate) => {
                const normalizedCandidate = normalizeHeaderKey(candidate);
                return normalizedCandidate === normalizedKey || (normalizedKey.length >= 4 && normalizedCandidate.startsWith(normalizedKey));
              });
              if (found !== undefined && row[found] !== undefined && row[found] !== "") return row[found];
            }
            return undefined;
          };
          const latitude = parseSpreadsheetNumber(getValue("latitude", "широта", "lat"));
          const longitude = parseSpreadsheetNumber(getValue("longitude", "довгота", "lon", "lng"));
          const sourceRegion = String(getValue("residence_region", "регіон місцезнаходження") ?? "").toLocaleLowerCase("uk-UA");
          if (registry.id === "healthcare" && isNszuDivisionsFile && !sourceRegion.includes("терноп")) return;
          const values: Record<string, string | number> = {};
          registry.fields.forEach((field) => {
            const raw = getValue(field.key, field.label);
            if (raw === undefined || raw === null || raw === "") return;
            if (field.type === "number") {
              const numeric = parseSpreadsheetNumber(raw);
              if (Number.isFinite(numeric)) values[field.key] = numeric;
            } else if (field.type === "date") {
              values[field.key] = normalizeDate(raw);
            } else {
              values[field.key] = String(raw).trim();
            }
          });
          if (registry.id === "healthcare" && isNszuDivisionsFile) {
            values.source = "НСЗУ — pmg_legal_entity_divisions_info.csv";
          }
          const validCoordinates = Number.isFinite(latitude) && Number.isFinite(longitude) && latitude >= -90 && latitude <= 90 && longitude >= -180 && longitude <= 180;
          const hasRequired = registry.fields.filter((field) => field.required).every((field) => values[field.key] !== undefined && String(values[field.key]).trim() !== "");
          if (!validCoordinates || !hasRequired) { sheetInvalidRows += 1; return; }
          const fallbackId = registry.id + ":" + String(values.object_id ?? values.name ?? "record") + ":" + index;
          stagedRecords.push({
            id: String(getValue("record_id", "recordId", "_record_id") ?? fallbackId),
            registryId: registry.id,
            values,
            latitude,
            longitude,
            createdAt: String(getValue("created_at", "createdAt") ?? new Date().toISOString())
          });
        });
        if (sheetInvalidRows > 0) {
          invalidRegistryRows += sheetInvalidRows;
          if (!(registry.id === "healthcare" && isNszuDivisionsFile)) return;
        }
        if (stagedRecords.length === 0) return;
        replacedRegistryIds.push(registry.id);
        importedRegistryRecords.push(...stagedRecords);
      });
      if (replacedRegistryIds.length) {
        setRegistryRecords((current) => [
          ...current.filter((record) => !replacedRegistryIds.includes(record.registryId)),
          ...importedRegistryRecords
        ]);
      }

      let importedLayers = 0;
      let invalidLayers = 0;
      const geoRows = readRows(GEO_LAYERS_SHEET, [["layer_id"], ["layer_name"], ["geojson_chunk"]], ["Географічні шари"]);
      if (geoRows !== null) {
        const grouped = new Map<string, { meta: Record<string, unknown>; chunks: Array<{ index: number; value: string }> }>();
        geoRows.forEach((row, index) => {
          const layerId = String(row.layer_id ?? row.id ?? "").trim() || ("excel-layer-" + index);
          const chunk = String(row.geojson_chunk ?? row.geojson ?? row.feature_collection ?? "").trim();
          const group = grouped.get(layerId) ?? { meta: row, chunks: [] };
          if (chunk) group.chunks.push({ index: Number(row.chunk_index ?? index), value: chunk });
          grouped.set(layerId, group);
        });
        const layers: GISLayer[] = [];
        grouped.forEach((group, layerId) => {
          try {
            const jsonText = group.chunks.sort((x, y) => x.index - y.index).map((item) => item.value).join("");
            const data = JSON.parse(jsonText);
            if (!data || data.type !== "FeatureCollection" || !Array.isArray(data.features)) throw new Error("Invalid FeatureCollection");
            layers.push({
              id: layerId,
              name: String(group.meta.layer_name ?? group.meta.name ?? "Географічний шар"),
              category: String(group.meta.category ?? "Імпортовані дані"),
              visible: !["false", "0", "ні", "no"].includes(String(group.meta.visible ?? "true").toLocaleLowerCase("uk-UA")),
              data,
              source: String(group.meta.source ?? "Імпорт із Excel"),
              importedAt: String(group.meta.imported_at ?? group.meta.importedAt ?? new Date().toISOString())
            });
          } catch { invalidLayers += 1; }
        });
        if (invalidLayers === 0) {
          setGisLayers(layers);
          importedLayers = layers.length;
        }
      }

      if (caseRows === null && replacedRegistryIds.length === 0 && geoRows === null) {
        setMessage("У книзі не знайдено аркушів, сумісних з імпортом EpiGIS.");
        return;
      }
      setSelectedDiseases([]);
      setSelectedHromadas([]);
      setSelected(null);
      const caseSummary = casesUpdated ? " Випадки: " + caseCount + " (без координат: " + noCoordinateCases + ", відхилено: " + invalidCases + ")." : " Випадки не змінювалися.";
      const registrySummary = replacedRegistryIds.length ? " Предметні реєстри: " + importedRegistryRecords.length + " записів у " + replacedRegistryIds.length + " аркушах." : " Предметні реєстри не змінювалися.";
      const layerSummary = geoRows === null || geoRows.length === 0
        ? " Геошари не змінювалися."
        : invalidLayers > 0
          ? " Геошари не оновлено через помилки GeoJSON; попередні шари збережено."
          : " Геошари: " + importedLayers + " шарів.";
      setMessage("Імпорт завершено." + caseSummary + registrySummary + layerSummary +
        (invalidRegistryRows || invalidLayers ? " Некоректні рядки реєстрів/геошарів: " + invalidRegistryRows + "/" + invalidLayers + "." : "") +
        caseImportWarning + " Дані збережено в локальному сховищі цього браузера.");
    } catch {
      setMessage("Не вдалося прочитати книгу. Перевір формат і заголовки; підтримуються книги EpiGIS, CSV НСЗУ з місцями надання послуг та таблиці випадків.");
    }
  };

  const importGeoJSON = async (file?: File) => {
    if (!file) return;
    try {
      const parsed = JSON.parse(await file.text());
      const data = parsed?.type === "FeatureCollection" ? parsed :
        parsed?.type === "Feature" ? { type: "FeatureCollection", features: [parsed] } : null;
      if (!data || !Array.isArray(data.features) || !data.features.length) {
        setLayerMessage("Файл має містити GeoJSON Feature або FeatureCollection з геометріями.");
        return;
      }
      const supported = data.features.filter((feature: any) =>
        feature?.type === "Feature" && feature.geometry &&
        ["Point", "MultiPoint", "LineString", "MultiLineString", "Polygon", "MultiPolygon"].includes(feature.geometry.type) &&
        hasValidGeoJSONCoordinates(feature.geometry.coordinates));
      if (!supported.length) {
        setLayerMessage("У файлі не знайдено підтримуваних геометрій: точки, лінії або полігони.");
        return;
      }
      const layer: GISLayer = {
        id: `layer-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
        name: file.name.replace(/\.(geojson|json)$/i, ""),
        category: activeRegistryId === "cases" ? CASE_REGISTRY.label : (REGISTRIES.find((registry) => registry.id === activeRegistryId)?.label ?? "Імпортовані дані"),
        visible: true,
        data: { type: "FeatureCollection", features: supported },
        source: file.name,
        importedAt: new Date().toISOString()
      };
      setGisLayers((current) => [...current, layer]);
      setLayerMessage(`Додано шар «${layer.name}»: ${supported.length} об'єктів. Відхилено некоректних або непідтримуваних: ${data.features.length - supported.length}.`);
    } catch {
      setLayerMessage("Не вдалося прочитати GeoJSON. Перевір структуру та синтаксис JSON.");
    } finally {
      if (layerFileRef.current) layerFileRef.current.value = "";
    }
  };

  const activeRegistry = activeRegistryId === "cases" ? CASE_REGISTRY : (REGISTRIES.find((registry) => registry.id === activeRegistryId) ?? REGISTRIES[0]);
  const activeRegistryRecords = registryRecords.filter((record) => record.registryId === activeRegistryId);
  const registryFilterDefinition = selectedRegistryIds.length === 1 ? REGISTRIES.find((registry) => registry.id === selectedRegistryIds[0]) : undefined;
  const registryTypeField = selectedRegistryIds.length !== 1 ? undefined : selectedRegistryIds[0] === "cases"
    ? { key: "disease", label: "Нозологія" }
    : registryFilterDefinition?.fields.find((field) => field.key.endsWith("_type") || ["disease", "organism", "life_stage", "risk_group"].includes(field.key));
  const registryTypeOptions = !registryTypeField ? [] : selectedRegistryIds[0] === "cases"
    ? Array.from(new Set(cases.map((item) => item.disease))).sort()
    : Array.from(new Set(registryRecords.filter((record) => record.registryId === selectedRegistryIds[0]).map((record) => String(record.values[registryTypeField.key] ?? "")).filter(Boolean))).sort();
  const registryCommunityOptions = ["Усі громади", ...Array.from(new Set([
    ...registryRecords.map((record) => String(record.values.hromada ?? "")),
    ...cases.map((item) => item.hromada)
  ].filter(Boolean))).sort()];
  const filteredRegistryRecords = registryRecords.filter((record) => {
    if (!selectedRegistryIds.includes(record.registryId)) return false;
    if (registryCommunityFilter !== "Усі громади" && String(record.values.hromada ?? "") !== registryCommunityFilter) return false;
    if (registryTypeFilter !== "Усі типи" && registryTypeField && record.registryId === selectedRegistryIds[0] && String(record.values[registryTypeField.key] ?? "") !== registryTypeFilter) return false;
    const dateValue = String(record.values.detected_at ?? record.values.sample_date ?? record.values.observed_at ?? record.values.updated_at ?? "").slice(0, 10);
    if (registryFromDate && (!dateValue || dateValue < registryFromDate)) return false;
    if (registryToDate && (!dateValue || dateValue > registryToDate)) return false;
    const registryName = REGISTRIES.find((registry) => registry.id === record.registryId)?.label ?? "";
    const searchable = [registryName, ...Object.values(record.values).map((value) => String(value ?? ""))].join(" ").toLocaleLowerCase("uk-UA");
    return !registrySearch || searchable.includes(registrySearch.toLocaleLowerCase("uk-UA"));
  });
  const openNewRegistryRecord = () => { setEditingRegistryRecordId(""); setRegistryFormValues({ updated_at: new Date().toISOString().slice(0, 10) }); setRegistryFormOpen(true); };
  const openEditRegistryRecord = (record: RegistryRecord) => { setEditingRegistryRecordId(record.id); setRegistryFormValues({ ...Object.fromEntries(Object.entries(record.values).map(([key, value]) => [key, String(value)])), latitude: String(record.latitude), longitude: String(record.longitude) }); setRegistryFormOpen(true); };
  const saveRegistryRecord = (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    const latitude = Number(registryFormValues.latitude); const longitude = Number(registryFormValues.longitude);
    if (!registryFormValues.name?.trim() || !Number.isFinite(latitude) || !Number.isFinite(longitude) || latitude < -90 || latitude > 90 || longitude < -180 || longitude > 180) { setLayerMessage("Потрібні назва об'єкта та коректні координати."); return; }
    const values: Record<string, string | number> = {};
    activeRegistry.fields.forEach((field) => { const value = registryFormValues[field.key]?.trim() ?? ""; if (value !== "") values[field.key] = field.type === "number" ? Number(value) : value; });
    if (activeRegistry.fields.some((field) => field.required && (values[field.key] === undefined || values[field.key] === ""))) { setLayerMessage("Заповни всі обов'язкові поля реєстру."); return; }
    const record: RegistryRecord = { id: editingRegistryRecordId || `reg-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`, registryId: activeRegistryId, values, latitude, longitude, createdAt: new Date().toISOString() };
    setRegistryRecords((current) => editingRegistryRecordId ? current.map((item) => item.id === editingRegistryRecordId ? record : item) : [...current, record]);
    setRegistryFormOpen(false); setEditingRegistryRecordId(""); setRegistryFormValues({});
    setLayerMessage(editingRegistryRecordId ? "Запис реєстру оновлено." : "Новий запис додано до реєстру.");
  };
  const deleteRegistryRecord = (id: string) => { if (!window.confirm("Видалити цей запис із локального реєстру?")) return; setRegistryRecords((current) => current.filter((record) => record.id !== id)); };

  const toggleLayer = (id: string) => setGisLayers((current) => current.map((layer) => layer.id === id ? { ...layer, visible: !layer.visible } : layer));
  const removeLayer = (id: string) => { setGisLayers((current) => current.filter((layer) => layer.id !== id)); if (fitLayerId === id) setFitLayerId(""); };
  const renameLayer = (layer: GISLayer) => {
    const proposed = window.prompt("Нова назва географічного шару:", layer.name);
    const name = proposed?.trim();
    if (!name || name === layer.name) return;
    setGisLayers((current) => current.map((item) => item.id === layer.id ? { ...item, name } : item));
    setLayerMessage(`Шар перейменовано на «${name}».`);
  };
  const exportGeoJSONLayer = (layer: GISLayer) => {
    const blob = new Blob([JSON.stringify(layer.data, null, 2)], { type: "application/geo+json;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `${layer.name.replace(/[^\p{L}\p{N}_-]+/gu, "_") || "gis-layer"}.geojson`;
    link.click();
    URL.revokeObjectURL(url);
  };



  const addMapFeature = (geometry: any, name: string, category: string) => {
    const feature = { type: "Feature", properties: { name, category, createdAt: new Date().toISOString(), source: "Створено на карті" }, geometry };
    const layer: GISLayer = {
      id: `draw-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
      name, category, visible: true,
      data: { type: "FeatureCollection", features: [feature] },
      source: "Ручне креслення на карті", importedAt: new Date().toISOString()
    };
    setGisLayers((current) => [...current, layer]);
    setLayerMessage(`Створено шар «${name}».`);
  };

  const handleMapDrawClick = (lat: number, lng: number) => {
    if (drawMode === "point") {
      addMapFeature({ type: "Point", coordinates: [lng, lat] }, "Нова точка", "Ручні об'єкти");
      setDrawMode("none");
      return;
    }
    if (drawMode === "line" || drawMode === "polygon") {
      setDrawPoints((current) => [...current, [lat, lng]]);
    }
  };

  const finishDrawing = () => {
    if (drawMode === "line" && drawPoints.length >= 2) {
      addMapFeature({ type: "LineString", coordinates: drawPoints.map(([lat, lng]) => [lng, lat]) }, "Нова лінія", "Ручні об'єкти");
    } else if (drawMode === "polygon" && drawPoints.length >= 3) {
      const ring = drawPoints.map(([lat, lng]) => [lng, lat]);
      ring.push(ring[0]);
      addMapFeature({ type: "Polygon", coordinates: [ring] }, "Новий полігон", "Ручні об'єкти");
    } else {
      setLayerMessage("Для лінії потрібно щонайменше 2 вершини, для полігона — 3.");
      return;
    }
    setDrawPoints([]);
    setDrawMode("none");
  };

  const cancelDrawing = () => { setDrawPoints([]); setDrawMode("none"); };

  const createBuffer = () => {
    const source = gisLayers.find((layer) => layer.id === bufferLayerId);
    const meters = Number(bufferMeters);
    if (!source) { setLayerMessage("Спочатку вибери шар для буферизації."); return; }
    if (!Number.isFinite(meters) || meters <= 0 || meters > 100000) { setLayerMessage("Вкажи відстань від 1 до 100 000 метрів."); return; }
    try {
      const result = turfBuffer(source.data, meters / 1000, { units: "kilometers", steps: 32 });
      if (!result) { setLayerMessage("Не вдалося побудувати буфер для вибраного шару."); return; }
      const resultFeatures = result.type === "FeatureCollection" ? result.features : [result];
      if (!resultFeatures.length) { setLayerMessage("Не вдалося побудувати буфер для вибраного шару."); return; }
      const layer: GISLayer = {
        id: `buffer-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
        name: `Буфер ${meters} м — ${source.name}`,
        category: "Буферні зони", visible: true, data: { type: "FeatureCollection", features: resultFeatures },
        source: `Буферизація шару «${source.name}»`, importedAt: new Date().toISOString()
      };
      setGisLayers((current) => [...current, layer]);
      setLayerMessage(`Побудовано буфер ${meters} м для ${resultFeatures.length} об'єктів шару «${source.name}».`);
    } catch {
      setLayerMessage("Помилка буферизації. Перевір геометрії у вихідному GeoJSON.");
    }
  };

  const runIntersection = () => {
    const zoneLayer = gisLayers.find((layer) => layer.id === intersectionZoneId);
    const targetLayer = gisLayers.find((layer) => layer.id === intersectionTargetId);
    if (!zoneLayer || !targetLayer) { setLayerMessage("Вибери шар зоною аналізу та шар об'єктів для перевірки."); return; }
    if (zoneLayer.id === targetLayer.id) { setLayerMessage("Для аналізу потрібно вибрати два різні шари."); return; }
    const zones = zoneLayer.data?.features ?? [];
    const targets = targetLayer.data?.features ?? [];
    if (!zones.length || !targets.length) { setLayerMessage("Один із вибраних шарів не містить об'єктів."); return; }
    try {
      const matches = targets.flatMap((feature: any, index: number) => {
        if (!feature?.geometry) return [];
        const intersectingZones = zones.filter((zone: any) => zone?.geometry && booleanIntersects(feature, zone));
        if (!intersectingZones.length) return [];
        return [{
          ...feature,
          properties: {
            ...(feature.properties ?? {}),
            _epigis_source_layer: targetLayer.name,
            _epigis_intersection_count: intersectingZones.length,
            _epigis_zone_layer: zoneLayer.name,
            _epigis_source_index: index
          }
        }];
      });
      if (!matches.length) {
        setLayerMessage(`Перетинів не виявлено: ${targets.length} об'єктів перевірено щодо шару «${zoneLayer.name}».`);
        return;
      }
      const resultLayer: GISLayer = {
        id: `intersection-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
        name: `Перетини: ${targetLayer.name} / ${zoneLayer.name}`,
        category: "Результати просторового аналізу",
        visible: true,
        data: { type: "FeatureCollection", features: matches },
        source: `Просторовий перетин «${targetLayer.name}» із «${zoneLayer.name}»`,
        importedAt: new Date().toISOString()
      };
      setGisLayers((current) => [...current, resultLayer]);
      setIntersectionResultLayerId(resultLayer.id);
      setLayerMessage(`Знайдено ${matches.length} об'єктів із перетином із зоною. Перевірено об'єктів: ${targets.length}; геометрій зони: ${zones.length}. Результат додано окремим шаром.`);
    } catch {
      setLayerMessage("Не вдалося виконати перетин. Перевір валідність геометрій у двох GeoJSON-шарах.");
    }
  };

  const exportMapImage = async () => {
    const target = mapExportRef.current;
    if (!target || mapExporting) return;
    setMapExporting(true);
    try {
      await new Promise((resolve) => setTimeout(resolve, 250));
      const canvas = await html2canvas(target, {
        backgroundColor: "#ffffff",
        scale: 2,
        useCORS: true,
        allowTaint: false,
        logging: false,
        imageTimeout: 20000
      });
      const stamp = new Date().toISOString().slice(0, 10);
      if (mapExportFormat === "pdf") {
        const image = canvas.toDataURL("image/png");
        const pdf = new jsPDF({ orientation: canvas.width >= canvas.height ? "landscape" : "portrait", unit: "mm", format: "a4" });
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        const ratio = Math.min(pageWidth / canvas.width, pageHeight / canvas.height);
        const width = canvas.width * ratio;
        const height = canvas.height * ratio;
        pdf.addImage(image, "PNG", (pageWidth - width) / 2, (pageHeight - height) / 2, width, height);
        pdf.save(`ternopil-epigis-map-${stamp}.pdf`);
      } else {
        const mime = mapExportFormat === "jpeg" ? "image/jpeg" : "image/png";
        const dataUrl = canvas.toDataURL(mime, mapExportFormat === "jpeg" ? 0.92 : undefined);
        const link = document.createElement("a");
        link.href = dataUrl;
        link.download = `ternopil-epigis-map-${stamp}.${mapExportFormat === "jpeg" ? "jpg" : "png"}`;
        link.click();
      }
    } catch (error) {
      console.error("Map export failed", error);
      setLayerMessage("Не вдалося створити зображення карти. Перевір доступність картографічних тайлів і повтори спробу.");
    } finally {
      setMapExporting(false);
    }
  };

  const exportCsv = () => {
    const header = TEMPLATE_HEADERS.join(",");
    const rows = filtered.map((item) => TEMPLATE_HEADERS.map((key) => {
      const value = (item as unknown as Record<string, unknown>)[key] ?? "";
      return `"${String(value).replace(/"/g, '""')}"`;
    }).join(","));
    const blob = new Blob(["\uFEFF" + [header, ...rows].join("\r\n")], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url; link.download = "ternopil-epigis-cases.csv"; link.click();
    URL.revokeObjectURL(url);
  };

  const downloadTemplate = () => {
    const workbook = XLSX.utils.book_new();
    const savedCases = cases.length && !cases.every((item) => item.id.startsWith("DEMO-"))
      ? cases.map((item) => TEMPLATE_HEADERS.map((key) => (item as unknown as Record<string, unknown>)[key] ?? ""))
      : [["CASE-0001", "Хвороба Лайма", "2026-05-12", "Тернопільська міська", "Тернопіль", 49.5535, 25.5948, "Підтверджено", 42, "Ж", "Приклад — замінити даними"]];
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([TEMPLATE_HEADERS, ...savedCases]), "Cases");

    REGISTRIES.forEach((registry) => {
      const headers = ["record_id", ...registry.fields.map((field) => field.key), "latitude", "longitude", "created_at"];
      const rows = registryRecords.filter((record) => record.registryId === registry.id).map((record) =>
        headers.map((key) => key === "record_id" ? record.id : key === "latitude" ? record.latitude : key === "longitude" ? record.longitude : key === "created_at" ? record.createdAt : record.values[key] ?? "")
      );
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([headers, ...rows]), REGISTRY_SHEET_NAMES[registry.id]);
    });

    const geoHeaders = ["layer_id", "layer_name", "category", "visible", "source", "imported_at", "chunk_index", "chunk_count", "geojson_chunk"];
    const geoRows: unknown[][] = [];
    gisLayers.forEach((layer) => {
      const serialized = JSON.stringify(layer.data);
      const chunkSize = 30000;
      const chunks: string[] = [];
      for (let offset = 0; offset < serialized.length; offset += chunkSize) chunks.push(serialized.slice(offset, offset + chunkSize));
      if (!chunks.length) chunks.push(serialized);
      chunks.forEach((chunk, index) => geoRows.push([
        layer.id, layer.name, layer.category, layer.visible ? "TRUE" : "FALSE", layer.source, layer.importedAt, index + 1, chunks.length, chunk
      ]));
    });
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([geoHeaders, ...geoRows]), GEO_LAYERS_SHEET);

    const infoRows = [
      ["Аркуш", "Призначення", "Правила заповнення"],
      ["Cases", "Випадки інфекційної захворюваності", "Чинна структура колонок не змінена. Обов’язкові поля: disease, date, hromada, latitude, longitude."],
      ...REGISTRIES.map((registry) => [REGISTRY_SHEET_NAMES[registry.id], registry.label, "Один рядок — один об’єкт. Обов’язкові поля визначені у відповідному реєстрі; координати latitude/longitude обов’язкові."]),
      ["GeoLayers", "GeoJSON-шари, водогони, трубопроводи, межі, буфери, ручні об'єкти та результати просторового аналізу", "Службовий аркуш. GeoJSON зберігається фрагментами у колонці geojson_chunk; не змінюй цей аркуш вручну."],
      ["Оновлення", "Книга є робочим файловим сховищем", "Завантаж книгу, додай або відредагуй записи, потім імпортуй її в EpiGIS. Імпортуються лише ті набори даних, аркуші яких містять заповнені рядки; порожні аркуші не очищують сховище."],
      ["Збереження", "Локальне сховище браузера", "Після імпорту записи зберігаються в цьому браузері. Завантажуй актуальну книгу як резервну копію; файл можна використовувати для перенесення даних на інший пристрій."],
      ["Конфіденційність", "Обмеження використання", "Не завантажуй персональні чи медично ідентифіковані дані. Населення внось лише агреговано. Серверної синхронізації між користувачами у цій версії немає."]
    ];
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(infoRows), "Instructions");

    const dictionaryRows: unknown[][] = [
      ["Аркуш", "Поле Excel", "Назва поля", "Тип даних", "Обов’язкове", "Допустимі значення / пояснення"],
      ...TEMPLATE_HEADERS.map((key) => ["Cases", key, key, key === "latitude" || key === "longitude" || key === "age" ? "Число" : key === "date" ? "Дата YYYY-MM-DD" : "Текст", ["disease", "date", "hromada", "latitude", "longitude"].includes(key) ? "Так" : "Ні", "Зберігається чинна структура аркуша Cases"]),
      ...REGISTRIES.flatMap((registry) => [
        ...registry.fields.map((field) => [REGISTRY_SHEET_NAMES[registry.id], field.key, field.label, field.type, field.required ? "Так" : "Ні", (field.options ?? []).join(" / ") ]),
        [REGISTRY_SHEET_NAMES[registry.id], "latitude", "Широта", "Число", "Так", "Десяткові градуси WGS84, від -90 до 90"],
        [REGISTRY_SHEET_NAMES[registry.id], "longitude", "Довгота", "Число", "Так", "Десяткові градуси WGS84, від -180 до 180"],
        [REGISTRY_SHEET_NAMES[registry.id], "record_id", "Службовий ID запису", "Текст", "Ні", "Можна залишити порожнім; система сформує ID автоматично"]
      ]),
      ["GeoLayers", "geojson_chunk", "Фрагмент GeoJSON", "Текст", "Так", "Службова структура; зберігає точки, лінії, полігони, буфери та інші геошари"],
      ["GeoLayers", "layer_id", "ID шару", "Текст", "Так", "Не змінюй вручну"],
      ["GeoLayers", "chunk_index / chunk_count", "Номер і кількість частин", "Число", "Так", "Не змінюй вручну"]
    ];
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(dictionaryRows), "DataDictionary");
    XLSX.writeFile(workbook, "ternopil-epigis-data-store.xlsx");
  };


  return localizeNode((
    <div className="app-shell">
      <header className="topbar">
        <div className="brand">
          <div className="brand-mark"><Activity size={22} strokeWidth={2.5} /></div>
          <div><div className="brand-name">Ternopil <span>EpiGIS</span></div><div className="brand-caption">ЕПІДЕМІОЛОГІЧНЕ КАРТОГРАФУВАННЯ</div></div>
        </div>
        <div className="top-actions"><button className={`button button-light language-switch ${language === "uk" ? "language-switch--uk" : "language-switch--en"}`} onClick={() => setLanguage((current) => current === "uk" ? "en" : "uk")} aria-label={language === "uk" ? "Мова інтерфейсу: українська. Перемкнути на English" : "Interface language: English. Switch to Ukrainian"} aria-pressed={language === "en"}><span className="language-switch__label">{language === "uk" ? "Мова" : "Language"}</span><span className="language-switch__option">УКР</span><span className="language-switch__separator">/</span><span className="language-switch__option">EN</span><span className="language-switch__active">{language === "uk" ? "● УКР" : "● EN"}</span></button><span className="environment"><span className="pulse-dot" /> Локальний режим</span><button className="button button-light" onClick={() => setShowUpload(!showUpload)}><Upload size={16} /> Імпорт даних</button></div>
      </header>

      <main className="main-content">
        <section className="page-heading">
          <div><div className="eyebrow">ТЕРНОПІЛЬСЬКА ОБЛАСТЬ / ЕПІДНАГЛЯД</div><h1>Карта епідеміологічної ситуації</h1><p>Просторовий аналіз випадків, часові фільтри та огляд територій.</p></div>
          <button className="button button-primary" onClick={exportCsv}><Download size={16} /> Експортувати вибірку</button>
        </section>

        {message && <div className="notice"><AlertTriangle size={16} /><span>{message}</span><button aria-label="Закрити повідомлення" onClick={() => setMessage("")}><X size={15} /></button></div>}

        {showUpload && <section className="upload-panel">
          <div className="upload-copy"><FileSpreadsheet size={22} /><div><strong>Імпорт книги EpiGIS або CSV НСЗУ</strong><p>Можна завантажити повну Excel-книгу EpiGIS або безпосередньо pmg_legal_entity_divisions_info.csv. Для CSV система відбере активні місця надання послуг Тернопільської області, зіставить поля та імпортує записи з валідними координатами.</p></div></div>
          <div className="upload-actions"><input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={(event) => importFile(event.target.files?.[0])} /><button className="button button-light" onClick={downloadTemplate}><Download size={15} /> Завантажити шаблон / сховище</button></div>
        </section>}

        <section className="metric-grid">
          <Metric icon={<Users size={18} />} label="Випадків у вибірці" value={filtered.length.toLocaleString("uk-UA")} note="За активними фільтрами" />
          <Metric icon={<ShieldCheck size={18} />} label="Підтверджено" value={confirmed.toLocaleString("uk-UA")} note={filtered.length ? `${Math.round(confirmed / filtered.length * 100)}% від вибірки` : "Немає записів"} />
          <Metric icon={<MapPin size={18} />} label="Громад із випадками" value={locations.toLocaleString("uk-UA")} note="У межах поточного відбору" />
          <Metric icon={<Activity size={18} />} label="Нозологій" value={activeDiseases.toLocaleString("uk-UA")} note="Різних діагнозів у вибірці" />
        </section>

        <section className="workspace">
          <aside className="sidebar">
            <div className="section-title"><Filter size={17} /><h2>Фільтри</h2><button className="text-button" onClick={() => {setSelectedDiseases([]);setSelectedHromadas([]);setFromDate("2026-01-01");setToDate("2026-12-31");setSearch("");setOnlyConfirmed(false);setSelectedRegistryIds(["cases"]);setRegistryCommunityFilter("Усі громади");setRegistryTypeFilter("Усі типи");setRegistrySearch("");setRegistryFromDate("");setRegistryToDate("");}}>Скинути</button></div>
            <label className="field-label">Пошук за записом або НП</label>
            <div className="search-field"><Search size={16} /><input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Діагноз, громада, населений пункт…" /></div>
            <label className="field-label">Нозології (можна обрати декілька)</label>
            <details className="multi-filter"><summary>{selectedDiseases.length ? `Обрано: ${selectedDiseases.length}` : "Усі нозології"} <span>▾</span></summary><div className="multi-filter-options"><button type="button" className="multi-filter-clear" onClick={() => setSelectedDiseases([])}>Очистити вибір</button>{diseases.map((item) => <label key={item}><input type="checkbox" checked={selectedDiseases.includes(item)} onChange={(e) => setSelectedDiseases((current) => e.target.checked ? [...current, item] : current.filter((value) => value !== item))} />{item}</label>)}</div></details>
            <label className="field-label">Територіальні громади (можна обрати декілька)</label>
            <details className="multi-filter"><summary>{selectedHromadas.length ? `Обрано: ${selectedHromadas.length}` : "Усі громади"} <span>▾</span></summary><div className="multi-filter-options"><button type="button" className="multi-filter-clear" onClick={() => setSelectedHromadas([])}>Очистити вибір</button>{hromadas.map((item) => <label key={item}><input type="checkbox" checked={selectedHromadas.includes(item)} onChange={(e) => setSelectedHromadas((current) => e.target.checked ? [...current, item] : current.filter((value) => value !== item))} />{item}</label>)}</div></details>
            <div className="date-fields"><div><label className="field-label"><CalendarDays size={13} /> Від дати</label><input type="date" value={fromDate} onChange={(e) => setFromDate(e.target.value)} /></div><div><label className="field-label"><CalendarDays size={13} /> До дати</label><input type="date" value={toDate} onChange={(e) => setToDate(e.target.value)} /></div></div>
            <label className="toggle-row"><input type="checkbox" checked={onlyConfirmed} onChange={(e) => setOnlyConfirmed(e.target.checked)} /><span className="toggle-ui" /><span>Лише підтверджені</span></label>
            <div className="sidebar-divider" />
            <div className="legend-title"><Filter size={15} /> Фільтри предметних реєстрів</div>
            <label className="field-label">Тип реєстру (можна обрати декілька)</label>
            <details className="multi-filter">
              <summary>{selectedRegistryIds.length ? `Обрано: ${selectedRegistryIds.length}` : "Реєстри не вибрано"} <span>▾</span></summary>
              <div className="multi-filter-options">
                <button type="button" className="multi-filter-clear" onClick={() => { setSelectedRegistryIds(["cases"]); setRegistryTypeFilter("Усі типи"); }}>За замовчуванням: реєстр випадків</button>
                <button type="button" className="multi-filter-clear" onClick={() => { setSelectedRegistryIds([]); setRegistryTypeFilter("Усі типи"); }}>Очистити вибір</button>
                <label><input type="checkbox" checked={selectedRegistryIds.includes("cases")} onChange={(event) => { setSelectedRegistryIds((current) => event.target.checked ? [...current, "cases"] : current.filter((id) => id !== "cases")); setRegistryTypeFilter("Усі типи"); }} />Реєстр випадків</label>
                {REGISTRIES.map((registry) => <label key={registry.id}><input type="checkbox" checked={selectedRegistryIds.includes(registry.id)} onChange={(event) => { setSelectedRegistryIds((current) => event.target.checked ? [...current, registry.id] : current.filter((id) => id !== registry.id)); setRegistryTypeFilter("Усі типи"); }} />{registry.label}</label>)}
              </div>
            </details>
            <label className="field-label">Пошук об’єкта / атрибута</label>
            <div className="search-field"><Search size={16} /><input value={registrySearch} onChange={(event) => setRegistrySearch(event.target.value)} placeholder="Назва, адреса, оператор, вид…" /></div>
            <label className="field-label">Територіальна громада</label>
            <div className="select-wrap"><select value={registryCommunityFilter} onChange={(event) => setRegistryCommunityFilter(event.target.value)}>{registryCommunityOptions.map((item) => <option key={item}>{item}</option>)}</select><ChevronDown size={15} /></div>
            {registryTypeField && selectedRegistryIds.length === 1 && <><label className="field-label">{registryTypeField.label}</label><div className="select-wrap"><select value={registryTypeFilter} onChange={(event) => setRegistryTypeFilter(event.target.value)}><option>Усі типи</option>{registryTypeOptions.map((item) => <option key={item}>{item}</option>)}</select><ChevronDown size={15} /></div></>}
            <div className="date-fields"><div><label className="field-label"><CalendarDays size={13} /> Від дати</label><input type="date" value={registryFromDate} onChange={(event) => setRegistryFromDate(event.target.value)} /></div><div><label className="field-label"><CalendarDays size={13} /> До дати</label><input type="date" value={registryToDate} onChange={(event) => setRegistryToDate(event.target.value)} /></div></div>
            <div className="layer-empty">Знайдено об’єктів реєстрів: {filteredRegistryRecords.length}</div>
            <div className="sidebar-divider" />
            <div className="legend-title"><Layers3 size={15} /> Умовні позначення</div>
            <div className="legend-row"><span className="legend-dot confirmed" /> Підтверджений випадок</div>
            <div className="legend-row"><span className="legend-dot investigating" /> Розслідується</div>
            <div className="sidebar-divider" />
            <div className="legend-title"><Layers3 size={15} /> Предметні реєстри</div>
            <label className="field-label">Тип реєстру для наступного GeoJSON-імпорту</label>
            <div className="select-wrap"><select value={activeRegistryId} onChange={(event) => { setActiveRegistryId(event.target.value); setRegistryFormOpen(false); setEditingRegistryRecordId(""); }}><option value="cases">Реєстр випадків</option>{REGISTRIES.map((registry) => <option key={registry.id} value={registry.id}>{registry.label}</option>)}</select><ChevronDown size={15} /></div>
            <p className="layer-empty">{activeRegistry.description}</p>
            <div className="registry-catalog"><div className="registry-catalog-item"><strong>Реєстр випадків</strong><span>{cases.length} записів · окремий набір даних</span></div>{REGISTRIES.map((registry) => <div className="registry-catalog-item" key={registry.id}><strong>{registry.label}</strong><span>{registry.fields.length} атрибутів · {registry.geometry === "mixed" ? "різні геометрії" : registry.geometry}</span></div>)}</div>
            <div className="sidebar-divider" />
            <div className="legend-title"><FileSpreadsheet size={15} /> Реєстр: {activeRegistry.label}</div>
            {activeRegistryId === "cases" ? <div className="layer-empty">Реєстр випадків доступний у таблиці «Реєстр випадків» нижче та через аркуш Cases в Excel-сховищі. Записи випадків не дублюються в предметних реєстрах.</div> : <><div className="registry-record-toolbar"><span>{activeRegistryRecords.length} записів</span><button className="button button-light" onClick={openNewRegistryRecord}>+ Додати запис</button></div>
            {registryFormOpen && <form className="registry-form" onSubmit={saveRegistryRecord}>
              <strong>{editingRegistryRecordId ? "Редагування запису" : "Новий запис"}</strong>
              {activeRegistry.fields.map((field) => <label className="field-label" key={field.key}>{field.label}{field.required ? " *" : ""}
                {field.type === "select" ? <select required={field.required} value={registryFormValues[field.key] ?? ""} onChange={(event) => setRegistryFormValues((current) => ({ ...current, [field.key]: event.target.value }))}><option value="">Вибрати…</option>{(field.options ?? []).map((option) => <option key={option} value={option}>{option}</option>)}</select> :
                <input required={field.required} type={field.type === "number" ? "number" : field.type === "date" ? "date" : "text"} value={registryFormValues[field.key] ?? ""} onChange={(event) => setRegistryFormValues((current) => ({ ...current, [field.key]: event.target.value }))} />}
              </label>)}
              <div className="registry-coordinate-grid"><label className="field-label">Широта *<input required type="number" step="any" min="-90" max="90" value={registryFormValues.latitude ?? ""} onChange={(event) => setRegistryFormValues((current) => ({ ...current, latitude: event.target.value }))} /></label><label className="field-label">Довгота *<input required type="number" step="any" min="-180" max="180" value={registryFormValues.longitude ?? ""} onChange={(event) => setRegistryFormValues((current) => ({ ...current, longitude: event.target.value }))} /></label></div>
              <div className="registry-form-actions"><button className="button button-primary" type="submit">Зберегти</button><button className="button button-light" type="button" onClick={() => setRegistryFormOpen(false)}>Скасувати</button></div>
            </form>}
            <div className="registry-record-list">{activeRegistryRecords.map((record) => <div className="registry-record" key={record.id}><div><strong>{String(record.values.name ?? record.values.object_id ?? "Об'єкт без назви")}</strong><span>{String(record.values.hromada ?? "Громаду не вказано")} · {record.latitude.toFixed(4)}, {record.longitude.toFixed(4)}</span></div><button className="layer-export" title="Редагувати запис" onClick={() => openEditRegistryRecord(record)}><Pencil size={13} /></button><button className="layer-remove" title="Видалити запис" onClick={() => deleteRegistryRecord(record.id)}><X size={13} /></button></div>)}</div>
            <p className="layer-empty">Записи зберігаються локально в цьому браузері. Координати використовуються для відображення точок на карті. Для ліній і полігонів імпортуй GeoJSON.</p></>}
          </aside>

          <section className="map-section">
            <div className="map-toolbar"><div><strong>Географічний огляд</strong><span>{filtered.length + filteredRegistryRecords.length} точкових записів на карті</span></div><div className="top-actions"><select className="map-export-select map-layer-select" aria-label="Підкладка карти" value={mapBaseLayer} onChange={(event) => setMapBaseLayer(event.target.value as "street" | "satellite")}><option value="street">Вулична карта</option><option value="satellite">Супутник</option></select><span className="map-chip"><span className="pulse-dot" /> Карта онлайн</span><select className="map-export-select" aria-label="Формат експорту карти" value={mapExportFormat} onChange={(event) => setMapExportFormat(event.target.value as "png" | "jpeg" | "pdf")}><option value="png">PNG</option><option value="jpeg">JPEG</option><option value="pdf">PDF</option></select><button className="button button-light map-export-button" onClick={exportMapImage} disabled={mapExporting}><Download size={15} /> {mapExporting ? "Створення…" : "Зберегти карту"}</button></div></div>
            <div className="map-wrap" ref={mapExportRef}>
              <MapContainer center={[49.45, 25.55]} zoom={8} zoomDelta={0.25} zoomSnap={0.25} wheelPxPerZoomLevel={120} scrollWheelZoom preferCanvas={true} className="leaflet-map">
                <DrawController mode={drawMode} onMapClick={handleMapDrawClick} />
                {drawMode === "line" && drawPoints.length > 0 && <Polyline positions={drawPoints} pathOptions={{ color: "#d97706", dashArray: "5 5" }} />}
                {drawMode === "polygon" && drawPoints.length > 0 && <Polygon positions={drawPoints} pathOptions={{ color: "#d97706", fillOpacity: 0.12, dashArray: "5 5" }} />}

                {mapBaseLayer === "satellite" ? <TileLayer attribution='&copy; Esri, Maxar, Earthstar Geographics' url="https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}" crossOrigin="anonymous" /> : <TileLayer attribution='&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors' url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png" crossOrigin="anonymous" />}
                {hromadaBoundaries && <GeoJSON
                  data={hromadaBoundaries}
                  style={(feature) => {
                    const boundaryName = String(feature?.properties?.ADM3_UA ?? "");
                    const isSelected = selectedHromadas.some((item) => boundaryName === item || boundaryName.replace(/ громада$/i, "") === item.replace(/ громада$/i, ""));
                    return {
                      color: isSelected ? "#0f766e" : "#315b73",
                      weight: isSelected ? 3 : 1.2,
                      opacity: 0.9,
                      fillColor: isSelected ? "#14b8a6" : "#78a6b8",
                      fillOpacity: isSelected ? 0.22 : 0.07
                    };
                  }}
                  onEachFeature={(feature, layer) => {
                    const name = String(feature?.properties?.ADM3_UA ?? "Територіальна громада");
                    const district = String(feature?.properties?.ADM2_UA ?? "");
                    layer.bindTooltip(district ? `${name} · ${district} район` : name, { sticky: true });
                    layer.on("click", () => {
                      const match = hromadas.find((item) => item !== "Усі громади" &&
                        (item.toLocaleLowerCase("uk-UA").includes(name.replace(/ громада$/i, "").toLocaleLowerCase("uk-UA")) ||
                         name.toLocaleLowerCase("uk-UA").includes(item.replace(/ громада$/i, "").toLocaleLowerCase("uk-UA"))));
                      if (match) setSelectedHromadas((current) => current.includes(match) ? current : [...current, match]);
                      else setSearch(name.replace(/ територіальна громада$/i, "").replace(/ громада$/i, ""));
                    });
                  }}
                />}
                {gisLayers.filter((layer) => layer.visible).map((layer, index) => <GeoJSON
                  key={layer.id}
                  data={layer.data}
                  style={() => ({
                    color: ["#176b55", "#2563a6", "#b7791f", "#7c3aed", "#be4b49"][index % 5],
                    weight: 2.5,
                    opacity: 0.9,
                    fillOpacity: 0.16
                  })}
                  pointToLayer={(_feature, latlng) => L.circleMarker(latlng, {
                    radius: 5, color: ["#176b55", "#2563a6", "#b7791f", "#7c3aed", "#be4b49"][index % 5],
                    fillColor: "#ffffff", fillOpacity: 0.9, weight: 2
                  })}
                  onEachFeature={(feature, layerObject) => {
                    const properties = feature?.properties ?? {};
                    const label = String(properties.name ?? properties.NAME ?? properties.НАЗВА ?? properties.id ?? "Об'єкт шару");
                    layerObject.bindTooltip(label, { sticky: true });
                  }}
                />)}
                <FitToCases cases={filtered} />
                <FitToLayer layer={gisLayers.find((layer) => layer.id === fitLayerId) ?? null} />
                <FacilityMarkerLayer records={filteredRegistryRecords.filter((record) => record.registryId === "healthcare")} language={language} />
                {filteredRegistryRecords.filter((record) => record.registryId !== "healthcare").map((record) => <CircleMarker key={record.id} center={[record.latitude, record.longitude]} radius={6} pathOptions={{ color: "#315f9a", fillColor: "#4d83bd", fillOpacity: 0.85, weight: 1.5 }}><Popup><strong>{String(record.values.name ?? "Об’єкт реєстру")}</strong><br />{REGISTRIES.find((registry) => registry.id === record.registryId)?.label}<br />{String(record.values.hromada ?? "")}</Popup></CircleMarker>)}
                {filtered.filter((item) => Number.isFinite(item.latitude) && Number.isFinite(item.longitude)).map((item) => <CircleMarker key={item.id} center={[item.latitude as number, item.longitude as number]} radius={selected?.id === item.id ? 10 : 7} pathOptions={{ color: item.status === "Підтверджено" ? "#126b58" : "#d28a24", fillColor: item.status === "Підтверджено" ? "#21866d" : "#e6a744", fillOpacity: 0.88, weight: selected?.id === item.id ? 3 : 1.5 }} eventHandlers={{ click: () => setSelected(item) }}><Popup><strong>{item.disease}</strong><br />{item.settlement || item.hromada}<br />{item.date}<br />{item.status}</Popup></CircleMarker>)}
              </MapContainer>
              {filtered.length === 0 && filteredRegistryRecords.length === 0 && <div className="empty-map">За вибраними параметрами записів не знайдено.</div>}
              <div className="map-disclaimer">Координати демонстраційних записів умовні.</div>
            </div>
            {filteredRegistryRecords.some((record) => record.registryId === "healthcare") && <div className="facility-map-legend"><span><i className="legend-hospital">✚</i> Лікарні</span><span><i className="legend-primary">⌂</i> Первинна допомога</span><span><i className="legend-dental">✦</i> Стоматології</span><span><i className="legend-emergency">✚</i> Екстрена допомога</span><span><i className="legend-laboratory">⚗</i> Лабораторії</span><span><i className="legend-clinic">●</i> Приватні центри</span></div>}
            {intersectionResultLayerId && gisLayers.some((layer) => layer.id === intersectionResultLayerId) && (() => {
              const resultLayer = gisLayers.find((layer) => layer.id === intersectionResultLayerId)!;
              const features = resultLayer.data?.features ?? [];
              return <section className="spatial-results">
                <div className="table-heading"><div><h2>Результати просторового аналізу</h2><p>{resultLayer.name} · {features.length} об’єктів</p></div><button className="text-button" onClick={() => setIntersectionResultLayerId("")}>Закрити таблицю</button></div>
                <div className="table-scroll"><table><thead><tr><th>№</th><th>Об’єкт</th><th>Тип геометрії</th><th>Кількість перетинів</th><th>Атрибути</th></tr></thead><tbody>
                  {features.map((feature: any, index: number) => {
                    const properties = feature?.properties ?? {};
                    const label = properties.name ?? properties.NAME ?? properties.НАЗВА ?? properties.id ?? `Об’єкт ${index + 1}`;
                    const attributes = Object.entries(properties).filter(([key]) => !key.startsWith("_epigis_")).slice(0, 5).map(([key, value]) => `${key}: ${typeof value === "object" ? JSON.stringify(value) : String(value ?? "")}`).join(" · ");
                    return <tr key={`${index}-${String(label)}`}><td>{index + 1}</td><td>{String(label)}</td><td>{feature?.geometry?.type ?? "—"}</td><td>{Number(properties._epigis_intersection_count ?? 0)}</td><td>{attributes || "Немає атрибутів"}</td></tr>;
                  })}
                </tbody></table></div>
                <p className="layer-empty">Таблиця відображає геометричні перетини, а не підтверджений епідеміологічний зв’язок. Для повного набору атрибутів експортуй GeoJSON відповідного шару.</p>
              </section>;
            })()}

            <div className="table-heading"><div><h2>Реєстр випадків</h2><p>Натисни на рядок, щоб переглянути деталі</p></div><span>{filtered.length} записів</span></div>
            <div className="table-scroll"><table><thead><tr><th>ID</th><th>Нозологія</th><th>Дата</th><th>Громада / НП</th><th>Статус</th></tr></thead><tbody>
              {filtered.slice(0, 100).map((item) => <tr key={item.id} className={selected?.id === item.id ? "selected-row" : ""} onClick={() => setSelected(item)}><td className="case-id">{item.id}</td><td>{item.disease}</td><td>{item.date}</td><td>{item.hromada}{item.settlement ? <small>{item.settlement}</small> : null}</td><td><span className={item.status === "Підтверджено" ? "status status-confirmed" : "status status-review"}>{item.status}</span></td></tr>)}
            </tbody></table>{filtered.length > 100 && <div className="table-foot">Показано 100 із {filtered.length} записів. Експорт містить усю вибірку.</div>}{!filtered.length && <div className="table-empty">Немає записів для відображення.</div>}</div>
            <div className="table-heading"><div><h2>Об’єкти предметних реєстрів</h2><p>За активними фільтрами реєстрів</p></div><span>{filteredRegistryRecords.length} записів</span></div>
            <div className="table-scroll"><table><thead><tr><th>Реєстр</th><th>Об’єкт</th><th>Громада / НП</th><th>Тип / характеристика</th><th>Дата</th></tr></thead><tbody>
              {filteredRegistryRecords.slice(0, 100).map((record) => {
                const registry = REGISTRIES.find((item) => item.id === record.registryId);
                const typeField = registry?.fields.find((field) => field.key.endsWith("_type") || ["disease", "organism", "life_stage", "risk_group"].includes(field.key));
                const dateValue = record.values.detected_at ?? record.values.sample_date ?? record.values.observed_at ?? record.values.updated_at ?? "";
                return <tr key={record.id}><td>{registry?.label ?? record.registryId}</td><td>{String(record.values.name ?? record.values.object_id ?? "—")}</td><td>{String(record.values.hromada ?? "—")}{record.values.settlement ? <small>{String(record.values.settlement)}</small> : null}</td><td>{String((typeField && record.values[typeField.key]) ?? "—")}</td><td>{String(dateValue || "—")}</td></tr>;
              })}
            </tbody></table>{filteredRegistryRecords.length > 100 && <div className="table-foot">Показано 100 із {filteredRegistryRecords.length} об’єктів. Повні дані містяться в Excel-сховищі.</div>}{!filteredRegistryRecords.length && <div className="table-empty">Об’єктів для відображення немає.</div>}</div>
          </section>

          <aside className="detail-panel">
            <div className="section-title"><Layers3 size={17} /><h2>Керування картою</h2></div>
            <div className="spatial-tools-panel">
            <div className="legend-title"><Layers3 size={15} /> Географічні шари <span className="layer-count">{gisLayers.length}</span></div>
            <button className="button button-light layer-import-button" onClick={() => layerFileRef.current?.click()}><Upload size={14} /> Додати GeoJSON</button>
            <input ref={layerFileRef} type="file" accept=".geojson,.json,application/geo+json,application/json" hidden onChange={(event) => importGeoJSON(event.target.files?.[0])} />
            {layerMessage && <div className="layer-message">{layerMessage}</div>}
            {gisLayers.length === 0 ? <p className="layer-empty">Поки немає власних шарів. Додай GeoJSON із точками, лініями або полігонами.</p> :
              <div className="layer-list">{gisLayers.map((layer) => <div className="layer-item" key={layer.id}>
                <label><input type="checkbox" checked={layer.visible} onChange={() => toggleLayer(layer.id)} /><span className="layer-name">{layer.name}</span></label>
                <span className="layer-meta">{layer.data.features.length} об'єктів</span>
                <button className="layer-export" title="Показати шар на карті" aria-label={`Масштабувати карту до шару ${layer.name}`} onClick={() => setFitLayerId(layer.id)}><LocateFixed size={13} /></button>
                <button className="layer-export" title="Перейменувати шар" aria-label={`Перейменувати шар ${layer.name}`} onClick={() => renameLayer(layer)}><Pencil size={13} /></button>
                <button className="layer-export" title="Експортувати GeoJSON" aria-label={`Експортувати шар ${layer.name}`} onClick={() => exportGeoJSONLayer(layer)}><Download size={13} /></button>
                <button className="layer-remove" title="Видалити шар" aria-label={`Видалити шар ${layer.name}`} onClick={() => removeLayer(layer.id)}><X size={13} /></button>
              </div>)}</div>}

            <div className="legend-title"><MapPin size={15} /> Креслення на карті</div>
            <div className="draw-tools">
              <button className={drawMode === "point" ? "draw-tool active" : "draw-tool"} onClick={() => { setDrawPoints([]); setDrawMode(drawMode === "point" ? "none" : "point"); }}>Точка</button>
              <button className={drawMode === "line" ? "draw-tool active" : "draw-tool"} onClick={() => { setDrawPoints([]); setDrawMode(drawMode === "line" ? "none" : "line"); }}>Лінія</button>
              <button className={drawMode === "polygon" ? "draw-tool active" : "draw-tool"} onClick={() => { setDrawPoints([]); setDrawMode(drawMode === "polygon" ? "none" : "polygon"); }}>Полігон</button>
            </div>
            {drawMode !== "none" && <div className="draw-status">
              <span>{drawMode === "point" ? "Клацни на карті, щоб поставити точку." : `Познач вершини на карті. Додано: ${drawPoints.length}.`}</span>
              {drawMode !== "point" && <button className="text-button" onClick={finishDrawing}>Завершити</button>}
              <button className="text-button" onClick={cancelDrawing}>Скасувати</button>
            </div>}
            <div className="legend-title"><Layers3 size={15} /> Буферний аналіз</div>
            <label className="field-label">Вихідний шар</label>
            <div className="select-wrap"><select value={bufferLayerId} onChange={(event) => setBufferLayerId(event.target.value)}><option value="">Вибрати шар…</option>{gisLayers.map((layer) => <option key={layer.id} value={layer.id}>{layer.name}</option>)}</select><ChevronDown size={15} /></div>
            <label className="field-label">Радіус буфера, метрів</label>
            <input className="buffer-input" type="number" min="1" max="100000" value={bufferMeters} onChange={(event) => setBufferMeters(event.target.value)} />
            <button className="button button-light layer-import-button" disabled={!gisLayers.length || !bufferLayerId} onClick={createBuffer}>Побудувати буфер</button>
            <div className="legend-title"><Layers3 size={15} /> Просторові перетини</div>
            <label className="field-label">Зона аналізу (наприклад, буфер)</label>
            <div className="select-wrap"><select value={intersectionZoneId} onChange={(event) => setIntersectionZoneId(event.target.value)}><option value="">Вибрати шар зоною…</option>{gisLayers.map((layer) => <option key={layer.id} value={layer.id}>{layer.name}</option>)}</select><ChevronDown size={15} /></div>
            <label className="field-label">Шар об'єктів для перевірки</label>
            <div className="select-wrap"><select value={intersectionTargetId} onChange={(event) => setIntersectionTargetId(event.target.value)}><option value="">Вибрати шар об'єктів…</option>{gisLayers.map((layer) => <option key={layer.id} value={layer.id}>{layer.name}</option>)}</select><ChevronDown size={15} /></div>
            <button className="button button-light layer-import-button" disabled={!intersectionZoneId || !intersectionTargetId} onClick={runIntersection}>Знайти об'єкти в зоні</button>
            <p className="layer-empty">Перевіряється геометричний перетин об'єктів двох шарів. Результат додається окремим шаром; це не є автоматичним епідеміологічним висновком.</p>
            <div className="privacy-note"><ShieldCheck size={16} /><span>Шари зберігаються локально у браузері. Вони не передаються на сервер і можуть бути недоступні на іншому пристрої.</span></div>
            </div>
            <div className="detail-divider" />
            <div className="section-title"><MapPin size={17} /><h2>Картка випадку</h2></div>
            {selected ? <><div className="detail-status"><span className={selected.status === "Підтверджено" ? "legend-dot confirmed" : "legend-dot investigating"} />{selected.status}</div><h3>{selected.disease}</h3><p className="detail-subtitle">{selected.id}</p><div className="detail-divider" />{[[ "Дата реєстрації", selected.date ],["Територія", selected.hromada],["Населений пункт", selected.settlement || "Не вказано"],["Вік", selected.age === undefined ? "Не вказано" : `${selected.age} років`],["Стать", selected.sex || "Не вказано"],["Широта", selected.latitude === undefined ? "Не вказано" : selected.latitude.toFixed(5)],["Довгота", selected.longitude === undefined ? "Не вказано" : selected.longitude.toFixed(5)]].map(([label, value]) => <div className="detail-field" key={label}><span>{label}</span><strong>{value}</strong></div>)}{selected.notes && <div className="detail-notes"><span>Примітки</span><p>{selected.notes}</p></div>}<button className="button button-light full-button" onClick={() => setSelected(null)}>Очистити вибір</button></> : <div className="detail-empty"><div className="detail-empty-icon"><MapPin size={22} /></div><strong>Випадок не вибрано</strong><p>Обери точку на карті або запис у реєстрі, щоб побачити деталі.</p></div>}
            <div className="alert-note"><AlertTriangle size={16} /><p><strong>Важливо</strong> Просторове скупчення точок саме по собі не підтверджує спалах. Потрібна епідеміологічна оцінка.</p></div>
          </aside>
        </section>
        <footer><span>Ternopil EpiGIS <b>·</b> MVP 0.1.0</span><span>Демонстраційні дані · Не для клінічних рішень</span><span className="developer-credit">Розробник: <a href="https://www.linkedin.com/in/taras-burda-410401318/" target="_blank" rel="noopener noreferrer">Taras Burda <span aria-hidden="true">↗</span></a> · <a href="mailto:burdataras@gmail.com">burdataras@gmail.com</a></span></footer>
      </main>
    </div>
  ), language);
}

function Metric({ icon, label, value, note }: { icon: React.ReactNode; label: string; value: string; note: string }) {
  return <div className="metric-card"><div className="metric-icon">{icon}</div><div className="metric-label">{label}</div><div className="metric-value">{value}</div><div className="metric-note">{note}</div></div>;
}
