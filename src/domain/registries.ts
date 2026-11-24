export type RegistryField = {
  key: string;
  label: string;
  type: "text" | "date" | "number" | "select";
  required?: boolean;
  options?: string[];
};

export type RegistryDefinition = {
  id: string;
  label: string;
  description: string;
  geometry: "Point" | "LineString" | "Polygon" | "mixed";
  fields: RegistryField[];
};

const commonLocation: RegistryField[] = [
  { key: "object_id", label: "Ідентифікатор", type: "text", required: true },
  { key: "name", label: "Назва об'єкта", type: "text", required: true },
  { key: "hromada", label: "Територіальна громада", type: "text" },
  { key: "settlement", label: "Населений пункт", type: "text" },
  { key: "address", label: "Адреса", type: "text" },
  { key: "source", label: "Джерело даних", type: "text" },
  { key: "updated_at", label: "Дата актуалізації", type: "date" }
];

export const REGISTRIES: RegistryDefinition[] = [
  {
    id: "healthcare",
    label: "Заклади охорони здоров'я",
    description: "Лікарні, амбулаторії, лабораторії, пункти екстреної допомоги.",
    geometry: "Point",
    fields: [...commonLocation, { key: "facility_type", label: "Тип закладу", type: "select", options: ["Лікарня", "Амбулаторія", "ЦПМСД", "Лабораторія", "Пункт екстреної допомоги", "Інше"] }, { key: "ownership", label: "Форма власності", type: "text" }]
  },
  {
    id: "water",
    label: "Водна інфраструктура",
    description: "Водозабори, водогони, каналізаційні мережі та очисні споруди.",
    geometry: "mixed",
    fields: [...commonLocation, { key: "asset_type", label: "Тип об'єкта", type: "select", options: ["Водозабір", "Водогін", "Каналізація", "Очисні споруди", "Інше"] }, { key: "operator", label: "Балансоутримувач / оператор", type: "text" }]
  },
  {
    id: "food",
    label: "Харчові об'єкти",
    description: "Підприємства харчування, виробництва, торгівлі та водопостачання харчових об'єктів.",
    geometry: "Point",
    fields: [...commonLocation, { key: "activity_type", label: "Вид діяльності", type: "text" }, { key: "risk_group", label: "Категорія ризику", type: "text" }]
  },
  {
    id: "natural",
    label: "Природні об'єкти",
    description: "Річки, водойми, ліси, заболочені території та інші природні об'єкти.",
    geometry: "mixed",
    fields: [...commonLocation, { key: "natural_type", label: "Тип природного об'єкта", type: "text" }]
  },
  {
    id: "outbreaks",
    label: "Епідеміологічні вогнища",
    description: "Вогнища, території розслідування та групові випадки; персональні дані не передбачені.",
    geometry: "mixed",
    fields: [...commonLocation, { key: "disease", label: "Нозологія", type: "text", required: true }, { key: "detected_at", label: "Дата виявлення", type: "date" }, { key: "status", label: "Статус", type: "select", options: ["Активне", "Під наглядом", "Закрите"] }]
  },
  {
    id: "bio-monitoring",
    label: "Біологічний моніторинг",
    description: "Місця відбору проб, кліщі, гризуни, переносники та результати лабораторних досліджень.",
    geometry: "Point",
    fields: [...commonLocation, { key: "sample_date", label: "Дата відбору", type: "date" }, { key: "organism", label: "Вид / таксон", type: "text" }, { key: "life_stage", label: "Стадія розвитку", type: "text" }, { key: "lab_result", label: "Результат дослідження", type: "text" }]
  },
  {
    id: "meteorology",
    label: "Метеорологія",
    description: "Метеостанції та спостереження: температура, опади, вологість, вітер.",
    geometry: "Point",
    fields: [...commonLocation, { key: "observed_at", label: "Дата і час спостереження", type: "date" }, { key: "temperature_c", label: "Температура, °C", type: "number" }, { key: "precipitation_mm", label: "Опади, мм", type: "number" }, { key: "humidity_pct", label: "Вологість, %", type: "number" }, { key: "wind_ms", label: "Швидкість вітру, м/с", type: "number" }]
  },
  {
    id: "population",
    label: "Населення",
    description: "Агрегована чисельність населення за населеними пунктами та громадами, без персональних записів.",
    geometry: "Point",
    fields: [...commonLocation, { key: "population_total", label: "Чисельність населення", type: "number", required: true }, { key: "reference_year", label: "Рік оцінки", type: "number" }]
  }
];

export const getRegistry = (id: string) => REGISTRIES.find((registry) => registry.id === id);
