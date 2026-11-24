import { useEffect, useMemo, useRef, useState } from "react";
import { CircleMarker, GeoJSON, MapContainer, Polygon, Polyline, Popup, TileLayer, useMap } from "react-leaflet";
import { buffer as turfBuffer } from "@turf/buffer";
import { booleanIntersects } from "@turf/boolean-intersects";
import L from "leaflet";
import { Activity, AlertTriangle, CalendarDays, ChevronDown, Download, FileSpreadsheet, Filter, Layers3, LocateFixed, MapPin, Pencil, Search, ShieldCheck, Upload, Users, X } from "lucide-react";
import * as XLSX from "xlsx";
import { REGISTRIES } from "./domain/registries";

type GISLayer = { id: string; name: string; category: string; visible: boolean; data: any; source: string; importedAt: string; };

type CaseRecord = {
  id: string;
  disease: string;
  date: string;
  hromada: string;
  settlement: string;
  latitude: number;
  longitude: number;
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

function FitToCases({ cases }: { cases: CaseRecord[] }) {
  const map = useMap();
  const first = useRef(true);
  if (first.current && cases.length) {
    map.setView([cases[0].latitude, cases[0].longitude], 8);
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

function normalizeDate(value: unknown): string {
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return [parsed.y, String(parsed.m).padStart(2, "0"), String(parsed.d).padStart(2, "0")].join("-");
  }
  const text = String(value ?? "").trim();
  if (/^\d{2}\.\d{2}\.\d{4}$/.test(text)) {
    const [d, m, y] = text.split(".");
    return `${y}-${m}-${d}`;
  }
  return text.slice(0, 10);
}

function toCase(row: Record<string, unknown>, index: number): CaseRecord | null {
  const get = (...keys: string[]) => {
    for (const key of keys) {
      const found = Object.keys(row).find((candidate) => candidate.trim().toLowerCase() === key.toLowerCase());
      if (found !== undefined && row[found] !== undefined && row[found] !== "") return row[found];
    }
    return undefined;
  };
  const latitude = Number(get("latitude", "lat", "широта"));
  const longitude = Number(get("longitude", "lon", "lng", "довгота"));
  const disease = String(get("disease", "нозологія", "діагноз", "final_diagnosis") ?? "").trim();
  const date = normalizeDate(get("date", "дата", "дата захворювання", "onset_date"));
  const hromada = String(get("hromada", "громада", "територія") ?? "").trim();
  if (!disease || !date || !hromada || !Number.isFinite(latitude) || !Number.isFinite(longitude)) return null;
  if (latitude < 44 || latitude > 53 || longitude < 20 || longitude > 42) return null;
  return {
    id: String(get("id", "номер", "case_id") ?? `IMP-${String(index + 1).padStart(4, "0")}`),
    disease, date, hromada,
    settlement: String(get("settlement", "населений пункт", "нп") ?? ""),
    latitude, longitude,
    status: String(get("status", "статус") ?? "Не визначено"),
    age: Number.isFinite(Number(get("age", "вік"))) ? Number(get("age", "вік")) : undefined,
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

export default function App() {
  const [cases, setCases] = useState<CaseRecord[]>(DEMO_CASES);
  const [gisLayers, setGisLayers] = useState<GISLayer[]>(() => {
    try { const saved = localStorage.getItem("ternopil-epigis-layers-v1"); return saved ? JSON.parse(saved) as GISLayer[] : []; } catch { return []; }
  });
  const [layerMessage, setLayerMessage] = useState("");
  const [activeRegistryId, setActiveRegistryId] = useState("healthcare");
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
  const [disease, setDisease] = useState("Усі нозології");
  const [hromada, setHromada] = useState("Усі громади");
  const [fromDate, setFromDate] = useState("2026-01-01");
  const [toDate, setToDate] = useState("2026-12-31");
  const [search, setSearch] = useState("");
  const [selected, setSelected] = useState<CaseRecord | null>(null);
  const [message, setMessage] = useState("Демонстраційні дані — не реальні випадки.");
  const [showUpload, setShowUpload] = useState(false);
  const [onlyConfirmed, setOnlyConfirmed] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);

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

  const diseases = useMemo(() => ["Усі нозології", ...Array.from(new Set(cases.map((item) => item.disease))).sort()], [cases]);
  const hromadas = useMemo(() => ["Усі громади", ...Array.from(new Set(cases.map((item) => item.hromada))).sort()], [cases]);
  const filtered = useMemo(() => cases.filter((item) =>
    (disease === "Усі нозології" || item.disease === disease) &&
    (hromada === "Усі громади" || item.hromada === hromada) &&
    (!fromDate || item.date >= fromDate) && (!toDate || item.date <= toDate) &&
    (!onlyConfirmed || item.status === "Підтверджено") &&
    (!search || [item.id, item.settlement, item.hromada, item.disease].join(" ").toLowerCase().includes(search.toLowerCase()))
  ), [cases, disease, hromada, fromDate, toDate, onlyConfirmed, search]);

  const confirmed = filtered.filter((item) => item.status === "Підтверджено").length;
  const locations = new Set(filtered.map((item) => item.hromada)).size;
  const activeDiseases = new Set(filtered.map((item) => item.disease)).size;

  const importFile = async (file?: File) => {
    if (!file) return;
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "" });
      const parsed = rows.map(toCase).filter((item): item is CaseRecord => item !== null);
      if (!parsed.length) {
        setMessage("Не знайдено валідних рядків. Перевір назви колонок та координати (latitude/longitude).");
        return;
      }
      setCases(parsed);
      setDisease("Усі нозології");
      setHromada("Усі громади");
      setSelected(null);
      setMessage(`Імпортовано ${parsed.length} записів із ${rows.length}. Відхилено: ${rows.length - parsed.length}.`);
    } catch {
      setMessage("Не вдалося прочитати файл. Спробуй XLSX, XLS або CSV.");
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
        category: REGISTRIES.find((registry) => registry.id === activeRegistryId)?.label ?? "Імпортовані дані",
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
    const worksheet = XLSX.utils.aoa_to_sheet([TEMPLATE_HEADERS, ["CASE-0001", "Хвороба Лайма", "2026-05-12", "Тернопільська міська", "Тернопіль", 49.5535, 25.5948, "Підтверджено", 42, "Ж", "Приклад — замінити даними"]]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Cases");
    XLSX.writeFile(workbook, "ternopil-epigis-import-template.xlsx");
  };



  return (
    <div className="app-shell">
      <header className="topbar">
        <div className="brand">
          <div className="brand-mark"><Activity size={22} strokeWidth={2.5} /></div>
          <div><div className="brand-name">Ternopil <span>EpiGIS</span></div><div className="brand-caption">ЕПІДЕМІОЛОГІЧНЕ КАРТОГРАФУВАННЯ</div></div>
        </div>
        <div className="top-actions"><span className="environment"><span className="pulse-dot" /> Локальний режим</span><button className="button button-light" onClick={() => setShowUpload(!showUpload)}><Upload size={16} /> Імпорт даних</button></div>
      </header>

      <main className="main-content">
        <section className="page-heading">
          <div><div className="eyebrow">ТЕРНОПІЛЬСЬКА ОБЛАСТЬ / ЕПІДНАГЛЯД</div><h1>Карта епідеміологічної ситуації</h1><p>Просторовий аналіз випадків, часові фільтри та огляд територій.</p></div>
          <button className="button button-primary" onClick={exportCsv}><Download size={16} /> Експортувати вибірку</button>
        </section>

        {message && <div className="notice"><AlertTriangle size={16} /><span>{message}</span><button aria-label="Закрити повідомлення" onClick={() => setMessage("")}><X size={15} /></button></div>}

        {showUpload && <section className="upload-panel">
          <div className="upload-copy"><FileSpreadsheet size={22} /><div><strong>Завантажити таблицю випадків</strong><p>Перший аркуш Excel або CSV. Обов’язкові поля: disease, date, hromada, latitude, longitude.</p></div></div>
          <div className="upload-actions"><input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={(event) => importFile(event.target.files?.[0])} /><button className="button button-light" onClick={downloadTemplate}><Download size={15} /> Шаблон Excel</button></div>
        </section>}

        <section className="metric-grid">
          <Metric icon={<Users size={18} />} label="Випадків у вибірці" value={filtered.length.toLocaleString("uk-UA")} note="За активними фільтрами" />
          <Metric icon={<ShieldCheck size={18} />} label="Підтверджено" value={confirmed.toLocaleString("uk-UA")} note={filtered.length ? `${Math.round(confirmed / filtered.length * 100)}% від вибірки` : "Немає записів"} />
          <Metric icon={<MapPin size={18} />} label="Громад із випадками" value={locations.toLocaleString("uk-UA")} note="У межах поточного відбору" />
          <Metric icon={<Activity size={18} />} label="Нозологій" value={activeDiseases.toLocaleString("uk-UA")} note="Різних діагнозів у вибірці" />
        </section>

        <section className="workspace">
          <aside className="sidebar">
            <div className="section-title"><Filter size={17} /><h2>Фільтри</h2><button className="text-button" onClick={() => {setDisease("Усі нозології");setHromada("Усі громади");setFromDate("2026-01-01");setToDate("2026-12-31");setSearch("");setOnlyConfirmed(false);}}>Скинути</button></div>
            <label className="field-label">Пошук за записом або НП</label>
            <div className="search-field"><Search size={16} /><input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Діагноз, громада, населений пункт…" /></div>
            <label className="field-label">Нозологія</label><div className="select-wrap"><select value={disease} onChange={(e) => setDisease(e.target.value)}>{diseases.map((item) => <option key={item}>{item}</option>)}</select><ChevronDown size={15} /></div>
            <label className="field-label">Територіальна громада</label><div className="select-wrap"><select value={hromada} onChange={(e) => setHromada(e.target.value)}>{hromadas.map((item) => <option key={item}>{item}</option>)}</select><ChevronDown size={15} /></div>
            <div className="date-fields"><div><label className="field-label"><CalendarDays size={13} /> Від дати</label><input type="date" value={fromDate} onChange={(e) => setFromDate(e.target.value)} /></div><div><label className="field-label"><CalendarDays size={13} /> До дати</label><input type="date" value={toDate} onChange={(e) => setToDate(e.target.value)} /></div></div>
            <label className="toggle-row"><input type="checkbox" checked={onlyConfirmed} onChange={(e) => setOnlyConfirmed(e.target.checked)} /><span className="toggle-ui" /><span>Лише підтверджені</span></label>
            <div className="sidebar-divider" />
            <div className="legend-title"><Layers3 size={15} /> Умовні позначення</div>
            <div className="legend-row"><span className="legend-dot confirmed" /> Підтверджений випадок</div>
            <div className="legend-row"><span className="legend-dot investigating" /> Розслідується</div>
            <div className="sidebar-divider" />
            <div className="legend-title"><Layers3 size={15} /> Предметні реєстри</div>
            <label className="field-label">Тип реєстру для наступного GeoJSON-імпорту</label>
            <div className="select-wrap"><select value={activeRegistryId} onChange={(event) => setActiveRegistryId(event.target.value)}>{REGISTRIES.map((registry) => <option key={registry.id} value={registry.id}>{registry.label}</option>)}</select><ChevronDown size={15} /></div>
            <p className="layer-empty">{REGISTRIES.find((registry) => registry.id === activeRegistryId)?.description}</p>
            <div className="registry-catalog">{REGISTRIES.map((registry) => <div className="registry-catalog-item" key={registry.id}><strong>{registry.label}</strong><span>{registry.fields.length} атрибутів · {registry.geometry === "mixed" ? "різні геометрії" : registry.geometry}</span></div>)}</div>
            <div className="sidebar-divider" />
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

            <div className="sidebar-divider" />
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
            <div className="sidebar-divider" />
            <div className="legend-title"><Layers3 size={15} /> Буферний аналіз</div>
            <label className="field-label">Вихідний шар</label>
            <div className="select-wrap"><select value={bufferLayerId} onChange={(event) => setBufferLayerId(event.target.value)}><option value="">Вибрати шар…</option>{gisLayers.map((layer) => <option key={layer.id} value={layer.id}>{layer.name}</option>)}</select><ChevronDown size={15} /></div>
            <label className="field-label">Радіус буфера, метрів</label>
            <input className="buffer-input" type="number" min="1" max="100000" value={bufferMeters} onChange={(event) => setBufferMeters(event.target.value)} />
            <button className="button button-light layer-import-button" disabled={!gisLayers.length || !bufferLayerId} onClick={createBuffer}>Побудувати буфер</button>
            <div className="sidebar-divider" />
            <div className="legend-title"><Layers3 size={15} /> Просторові перетини</div>
            <label className="field-label">Зона аналізу (наприклад, буфер)</label>
            <div className="select-wrap"><select value={intersectionZoneId} onChange={(event) => setIntersectionZoneId(event.target.value)}><option value="">Вибрати шар зоною…</option>{gisLayers.map((layer) => <option key={layer.id} value={layer.id}>{layer.name}</option>)}</select><ChevronDown size={15} /></div>
            <label className="field-label">Шар об'єктів для перевірки</label>
            <div className="select-wrap"><select value={intersectionTargetId} onChange={(event) => setIntersectionTargetId(event.target.value)}><option value="">Вибрати шар об'єктів…</option>{gisLayers.map((layer) => <option key={layer.id} value={layer.id}>{layer.name}</option>)}</select><ChevronDown size={15} /></div>
            <button className="button button-light layer-import-button" disabled={!intersectionZoneId || !intersectionTargetId} onClick={runIntersection}>Знайти об'єкти в зоні</button>
            <p className="layer-empty">Перевіряється геометричний перетин об'єктів двох шарів. Результат додається окремим шаром; це не є автоматичним епідеміологічним висновком.</p>
            <div className="privacy-note"><ShieldCheck size={16} /><span>Шари зберігаються локально у браузері. Вони не передаються на сервер і можуть бути недоступні на іншому пристрої.</span></div>
          </aside>

          <section className="map-section">
            <div className="map-toolbar"><div><strong>Географічний огляд</strong><span>{filtered.length} записів на карті</span></div><div className="top-actions"><span className="map-chip"><span className="pulse-dot" /> Карта онлайн</span></div></div>
            <div className="map-wrap">
              <MapContainer center={[49.45, 25.55]} zoom={8} scrollWheelZoom preferCanvas={true} className="leaflet-map">
                <DrawController mode={drawMode} onMapClick={handleMapDrawClick} />
                {drawMode === "line" && drawPoints.length > 0 && <Polyline positions={drawPoints} pathOptions={{ color: "#d97706", dashArray: "5 5" }} />}
                {drawMode === "polygon" && drawPoints.length > 0 && <Polygon positions={drawPoints} pathOptions={{ color: "#d97706", fillOpacity: 0.12, dashArray: "5 5" }} />}

                <TileLayer attribution='&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors' url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png" crossOrigin="anonymous" />
                {hromadaBoundaries && <GeoJSON
                  data={hromadaBoundaries}
                  style={(feature) => {
                    const boundaryName = String(feature?.properties?.ADM3_UA ?? "");
                    const isSelected = hromada !== "Усі громади" && (boundaryName === hromada || boundaryName.replace(/ громада$/i, "") === hromada.replace(/ громада$/i, ""));
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
                      if (match) setHromada(match);
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
                {filtered.map((item) => <CircleMarker key={item.id} center={[item.latitude, item.longitude]} radius={selected?.id === item.id ? 10 : 7} pathOptions={{ color: item.status === "Підтверджено" ? "#126b58" : "#d28a24", fillColor: item.status === "Підтверджено" ? "#21866d" : "#e6a744", fillOpacity: 0.88, weight: selected?.id === item.id ? 3 : 1.5 }} eventHandlers={{ click: () => setSelected(item) }}><Popup><strong>{item.disease}</strong><br />{item.settlement || item.hromada}<br />{item.date}<br />{item.status}</Popup></CircleMarker>)}
              </MapContainer>
              {filtered.length === 0 && <div className="empty-map">За вибраними параметрами записів не знайдено.</div>}
              <div className="map-disclaimer">Координати демонстраційних записів умовні.</div>
            </div>
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
          </section>

          <aside className="detail-panel">
            <div className="section-title"><MapPin size={17} /><h2>Картка випадку</h2></div>
            {selected ? <><div className="detail-status"><span className={selected.status === "Підтверджено" ? "legend-dot confirmed" : "legend-dot investigating"} />{selected.status}</div><h3>{selected.disease}</h3><p className="detail-subtitle">{selected.id}</p><div className="detail-divider" />{[[ "Дата реєстрації", selected.date ],["Територія", selected.hromada],["Населений пункт", selected.settlement || "Не вказано"],["Вік", selected.age === undefined ? "Не вказано" : `${selected.age} років`],["Стать", selected.sex || "Не вказано"],["Широта", selected.latitude.toFixed(5)],["Довгота", selected.longitude.toFixed(5)]].map(([label, value]) => <div className="detail-field" key={label}><span>{label}</span><strong>{value}</strong></div>)}{selected.notes && <div className="detail-notes"><span>Примітки</span><p>{selected.notes}</p></div>}<button className="button button-light full-button" onClick={() => setSelected(null)}>Очистити вибір</button></> : <div className="detail-empty"><div className="detail-empty-icon"><MapPin size={22} /></div><strong>Випадок не вибрано</strong><p>Обери точку на карті або запис у реєстрі, щоб побачити деталі.</p></div>}
            <div className="alert-note"><AlertTriangle size={16} /><p><strong>Важливо</strong> Просторове скупчення точок саме по собі не підтверджує спалах. Потрібна епідеміологічна оцінка.</p></div>
          </aside>
        </section>
        <footer><span>Ternopil EpiGIS <b>·</b> MVP 0.1.0</span><span>Демонстраційні дані · Не для клінічних рішень</span><span className="developer-credit">Розробник: <a href="https://www.linkedin.com/in/taras-burda-410401318/" target="_blank" rel="noopener noreferrer">Taras Burda <span aria-hidden="true">↗</span></a></span></footer>
      </main>
    </div>
  );
}

function Metric({ icon, label, value, note }: { icon: React.ReactNode; label: string; value: string; note: string }) {
  return <div className="metric-card"><div className="metric-icon">{icon}</div><div className="metric-label">{label}</div><div className="metric-value">{value}</div><div className="metric-note">{note}</div></div>;
}