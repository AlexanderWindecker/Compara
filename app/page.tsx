"use client";

import {
  ArrowDownToLine,
  ArrowDownUp,
  ArrowRight,
  ArrowUpRight,
  Award,
  BarChart3,
  BookmarkPlus,
  Check,
  ChevronDown,
  CircleHelp,
  Coffee,
  Download,
  FileSpreadsheet,
  Filter,
  Layers3,
  LoaderCircle,
  LockKeyhole,
  Moon,
  Plus,
  RefreshCw,
  Rows3,
  Search,
  Settings2,
  Sparkles,
  Table2,
  Trash2,
  Sun,
  Upload,
  UsersRound,
  X,
} from "lucide-react";
import * as XLSX from "xlsx";
import {
  ChangeEvent,
  DragEvent,
  FormEvent,
  Fragment,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";

type Cell = string | number | boolean | Date | null;
type DataRow = Record<string, Cell>;
type Grouping = "month" | "year" | "quarter" | "day";
type MetricMode = "sum" | "count" | "average";
type WorkbookData = {
  name: string;
  sheet: string;
  sheets: string[];
  buffer: ArrayBuffer;
  rows: DataRow[];
  headers: string[];
};
type ProductRequirement = {
  label: string;
  category: string;
  minimum: number;
};
type ProductSale = {
  quantity: number;
  meetsRequirement: boolean;
  usesSavedResult: boolean;
};
type ProductReportRow = {
  code: string;
  customer: string;
  seller: string;
  products: ProductSale[];
  meetsAll: boolean;
  missingProducts: string[];
  hasProductSales: boolean;
};
type ProductReport = {
  name: string;
  rows: ProductReportRow[];
  sellers: string[];
  transactionCount: number;
  formulaResultRows: number;
};
type ConsolidatedWorkbook = {
  name: string;
  headers: string[];
  rows: Cell[][];
  sheets: string[];
};
type ProductReportFilter = "all" | "complete" | "incomplete" | "no-sales";
type ComparisonSource = WorkbookData & {
  id: string;
  label: string;
  fieldMapping: Record<string, string>;
};
type FileComparisonMode = "periods" | "separate";
type PivotCell = {
  total: number;
  count: number;
};
type MetricValues = Map<string, PivotCell>;
type PivotRowValues = Map<string, MetricValues>;
type DimensionFilter = {
  field: string;
  values: string[];
};
type SavedComparison = {
  id: string;
  name: string;
  rowField: string;
  columnFields: string[];
  metricFields: string[];
  dateGrouping: Grouping;
  metricMode: MetricMode;
  filters: DimensionFilter[];
};

const SAVED_COMPARISONS_KEY = "compara-saved-comparisons";
const PRODUCT_REQUIREMENTS: ProductRequirement[] = [
  { label: "Café", category: "cafe", minimum: 6 },
  { label: "Edulco", category: "edulco", minimum: 1 },
  { label: "Azúcar", category: "azucar", minimum: 1 },
  { label: "Yerba", category: "yerba", minimum: 1 },
  { label: "Té", category: "te", minimum: 1 },
];

function isSavedComparison(value: unknown): value is SavedComparison {
  if (!value || typeof value !== "object") return false;
  const comparison = value as Partial<SavedComparison>;
  return (
    typeof comparison.id === "string" &&
    typeof comparison.name === "string" &&
    typeof comparison.rowField === "string" &&
    Array.isArray(comparison.columnFields) &&
    comparison.columnFields.every((field) => typeof field === "string") &&
    Array.isArray(comparison.metricFields) &&
    comparison.metricFields.every((field) => typeof field === "string") &&
    ["month", "year", "quarter", "day"].includes(comparison.dateGrouping ?? "") &&
    ["sum", "average", "count"].includes(comparison.metricMode ?? "") &&
    Array.isArray(comparison.filters) &&
    comparison.filters.every(
      (filter) =>
        filter &&
        typeof filter.field === "string" &&
        Array.isArray(filter.values) &&
        filter.values.every((entry) => typeof entry === "string"),
    )
  );
}

function isProductReportFilter(value: string): value is ProductReportFilter {
  return ["all", "complete", "incomplete", "no-sales"].includes(value);
}

const MONTHS = [
  "Enero",
  "Febrero",
  "Marzo",
  "Abril",
  "Mayo",
  "Junio",
  "Julio",
  "Agosto",
  "Septiembre",
  "Octubre",
  "Noviembre",
  "Diciembre",
];

const MONTH_LOOKUP: Record<string, number> = {
  enero: 0,
  febrero: 1,
  marzo: 2,
  abril: 3,
  mayo: 4,
  junio: 5,
  julio: 6,
  agosto: 7,
  septiembre: 8,
  setiembre: 8,
  octubre: 9,
  noviembre: 10,
  diciembre: 11,
};

const PALETTE = ["#356f51", "#dda957", "#83a88b", "#c8885c", "#697e67"];
const MAX_UPLOAD_BYTES = 50 * 1024 * 1024;
const MAX_EXCEL_ROWS = 1_048_576;
const MAX_EXCEL_COLUMNS = 16_384;
const CONSOLIDATION_CHUNK_SIZE = 10_000;

function cleanHeader(value: unknown, index: number) {
  const header = String(value ?? "").trim();
  return header || `Columna ${index + 1}`;
}

function normalizeText(value: string) {
  return value
    .trim()
    .toLocaleLowerCase("es")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function normalizeGroupingValue(value: string) {
  return value.trim().replace(/\s+/g, " ").toLocaleLowerCase("es");
}

function displayGroupingValue(value: string) {
  return value.trim().replace(/\s+/g, " ");
}

function findHeader(
  headers: string[],
  pattern: RegExp,
  fallback?: string,
): string {
  return headers.find((header) => pattern.test(normalizeText(header))) ?? fallback ?? "";
}

function isValidDate(value: unknown): value is Date {
  return value instanceof Date && !Number.isNaN(value.getTime());
}

function asDate(value: Cell): Date | null {
  if (isValidDate(value)) return value;
  if (typeof value === "number" && value > 20000 && value < 80000) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
  }
  if (typeof value === "string" && value.trim()) {
    const localDate = value.trim().match(
      /^(\d{1,2})[/. -](\d{1,2})[/. -](\d{4})(?:[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/,
    );
    if (localDate) {
      const [, day, month, year, hour = "0", minute = "0", second = "0"] = localDate;
      const parsed = new Date(
        Number(year),
        Number(month) - 1,
        Number(day),
        Number(hour),
        Number(minute),
        Number(second),
      );
      if (
        parsed.getFullYear() === Number(year) &&
        parsed.getMonth() === Number(month) - 1 &&
        parsed.getDate() === Number(day)
      ) {
        return parsed;
      }
      return null;
    }
    const parsed = new Date(value);
    if (!Number.isNaN(parsed.getTime())) return parsed;
    const month = MONTH_LOOKUP[normalizeText(value)];
    if (month !== undefined) return new Date(2025, month, 1);
  }
  return null;
}

function dateValue(date: Date, grouping: Grouping) {
  if (grouping === "year") return String(date.getFullYear());
  if (grouping === "quarter") return `T${Math.ceil((date.getMonth() + 1) / 3)} ${date.getFullYear()}`;
  if (grouping === "day")
    return `${String(date.getDate()).padStart(2, "0")}/${String(date.getMonth() + 1).padStart(2, "0")}/${date.getFullYear()}`;
  return `${MONTHS[date.getMonth()]} ${date.getFullYear()}`;
}

function dateSort(value: string) {
  const quarter = value.match(/^T(\d) (\d{4})$/);
  if (quarter) return Number(quarter[2]) * 10 + Number(quarter[1]);
  const month = value.match(/^(.+) (\d{4})$/);
  if (month) {
    const index = MONTH_LOOKUP[normalizeText(month[1])];
    if (index !== undefined) return Number(month[2]) * 100 + index;
  }
  const day = value.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (day) return Number(day[3]) * 10000 + Number(day[2]) * 100 + Number(day[1]);
  const year = Number(value);
  return Number.isFinite(year) ? year : Number.MAX_SAFE_INTEGER;
}

function cellText(value: Cell) {
  if (value === null || value === undefined || value === "") return "Sin dato";
  if (isValidDate(value)) return value.toLocaleDateString("es-AR");
  return String(value);
}

function pivotCellValue(cell: PivotCell | undefined, mode: MetricMode) {
  if (!cell) return 0;
  if (mode === "count") return cell.count;
  if (mode === "average") return cell.total / Math.max(1, cell.count);
  return cell.total;
}

function formatNumber(value: number, decimals = 0) {
  return new Intl.NumberFormat("es-AR", {
    maximumFractionDigits: decimals,
    minimumFractionDigits: decimals,
  }).format(value);
}

function formatMoney(value: number) {
  return new Intl.NumberFormat("es-AR", {
    style: "currency",
    currency: "ARS",
    maximumFractionDigits: 0,
  }).format(value);
}

function parseWorkbookSheet(
  workbook: XLSX.WorkBook,
  sheetName: string,
): Pick<WorkbookData, "rows" | "headers"> {
  const worksheet = workbook.Sheets[sheetName];
  const rawRows = XLSX.utils.sheet_to_json<unknown[]>(worksheet, {
    header: 1,
    defval: null,
    raw: true,
    blankrows: false,
  });
  const rawHeaders = rawRows[0] ?? [];
  const headers = rawHeaders.map(cleanHeader);
  const seen = new Map<string, number>();
  const uniqueHeaders = headers.map((header) => {
    const count = (seen.get(header) ?? 0) + 1;
    seen.set(header, count);
    return count > 1 ? `${header} (${count})` : header;
  });
  const rows = rawRows.slice(1).map((rawRow) => {
    const row: DataRow = {};
    uniqueHeaders.forEach((header, index) => {
      const value = rawRow[index];
      row[header] = value instanceof Date || typeof value === "string" || typeof value === "number" || typeof value === "boolean"
        ? value
        : null;
    });
    return row;
  });
  return { rows, headers: uniqueHeaders };
}

function findProductReportHeader(
  worksheet: XLSX.WorkSheet,
  includeSeller: boolean,
  preferReasonSocial: boolean,
) {
  const range = XLSX.utils.decode_range(worksheet["!ref"] ?? "A1");
  const lastHeaderRow = Math.min(range.e.r, range.s.r + 20);
  for (let row = range.s.r; row <= lastHeaderRow; row += 1) {
    const headers = Array.from(
      { length: range.e.c - range.s.c + 1 },
      (_, index) => {
        const column = range.s.c + index;
        const value = worksheet[XLSX.utils.encode_cell({ r: row, c: column })]?.v;
        return { column, value: normalizeText(String(value ?? "")) };
      },
    );
    const findColumn = (pattern: RegExp) =>
      headers.find(({ value }) => pattern.test(value))?.column;
    const customerColumn = preferReasonSocial
      ? findColumn(/razon\s*social|nombre\s*(del\s*)?cliente/) ??
        findColumn(/^clientes?$/)
      : findColumn(/^clientes?$/) ??
        findColumn(/razon\s*social|nombre\s*(del\s*)?cliente/);
    const sellerColumn = includeSeller
      ? findColumn(/^vendedora?$/)
      : undefined;
    const quantityColumn = includeSeller
      ? undefined
      : findColumn(/^cantidad$|unidades?/);
    const categoryColumn = includeSeller
      ? undefined
      : findColumn(/^categoria$|rubro/);
    if (
      customerColumn !== undefined &&
      (includeSeller
        ? sellerColumn !== undefined
        : quantityColumn !== undefined && categoryColumn !== undefined)
    ) {
      const codeColumn = headers.find(({ value }) =>
        /^codigo$|^cod\.?$/.test(value),
      )?.column;
      const productStatusColumns = PRODUCT_REQUIREMENTS.map((product) =>
        headers.find(
          ({ value }) =>
            value === product.category ||
            value.startsWith(`${product.category} `),
        )?.column,
      );
      return {
        headerRow: row,
        customerColumn,
        sellerColumn,
        quantityColumn,
        categoryColumn,
        codeColumn,
        productStatusColumns,
      };
    }
  }
  return null;
}

function readWorksheetValue(
  worksheet: XLSX.WorkSheet,
  row: number,
  column: number | undefined,
) {
  if (column === undefined) return null;
  return worksheet[XLSX.utils.encode_cell({ r: row, c: column })]?.v ?? null;
}

function parseQuantity(value: unknown, rowNumber: number) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "string") {
    let normalized = value.trim().replace(/\s/g, "");
    if (normalized.includes(",") && normalized.includes(".")) {
      normalized =
        normalized.lastIndexOf(",") > normalized.lastIndexOf(".")
          ? normalized.replace(/\./g, "").replace(",", ".")
          : normalized.replace(/,/g, "");
    } else if (normalized.includes(",")) {
      normalized = normalized.replace(",", ".");
    }
    const quantity = Number(normalized);
    if (Number.isFinite(quantity)) return quantity;
  }
  throw new Error(`La cantidad de la fila ${rowNumber} no es válida.`);
}

function savedFormulaResult(value: unknown): boolean | null {
  const result = String(value ?? "").trim();
  if (result === "✅" || ["si", "true", "1", "cumple"].includes(normalizeText(result))) {
    return true;
  }
  if (result === "❌" || ["no", "false", "0", "no cumple"].includes(normalizeText(result))) {
    return false;
  }
  return null;
}

function createProductReport(
  workbook: XLSX.WorkBook,
  fileName: string,
): ProductReport {
  const baseSheet = workbook.SheetNames
    .map((name) => workbook.Sheets[name])
    .find((sheet) => findProductReportHeader(sheet, true, false));
  const transactionSheet = workbook.SheetNames
    .map((name) => workbook.Sheets[name])
    .find((sheet) => findProductReportHeader(sheet, false, true));
  if (!baseSheet || !transactionSheet) {
    throw new Error(
      "El Excel debe incluir una hoja de clientes (CLIENTES y VENDEDOR) y otra de ventas (RAZÓN SOCIAL, CANTIDAD y CATEGORÍA).",
    );
  }
  const baseHeader = findProductReportHeader(baseSheet, true, false);
  const transactionHeader = findProductReportHeader(
    transactionSheet,
    false,
    true,
  );
  if (!baseHeader || !transactionHeader) {
    throw new Error("No pudimos reconocer los encabezados de las dos hojas.");
  }

  const baseRange = XLSX.utils.decode_range(baseSheet["!ref"] ?? "A1");
  const clients: {
    code: string;
    customer: string;
    seller: string;
    rowIndex: number;
  }[] = [];
  for (let row = baseHeader.headerRow + 1; row <= baseRange.e.r; row += 1) {
    const customer = String(
      readWorksheetValue(baseSheet, row, baseHeader.customerColumn) ?? "",
    ).trim();
    if (!customer) continue;
    const seller = String(
      readWorksheetValue(baseSheet, row, baseHeader.sellerColumn) ?? "",
    ).trim();
    const codeValue = readWorksheetValue(baseSheet, row, baseHeader.codeColumn);
    clients.push({
      code: codeValue === null ? "" : String(codeValue).trim(),
      customer,
      seller: seller || "Sin vendedor",
      rowIndex: row,
    });
  }
  if (clients.length === 0) {
    throw new Error("La hoja de clientes no contiene clientes para informar.");
  }

  const quantitiesByCustomer = new Map<string, number[]>();
  const customerColumnLetter = XLSX.utils.encode_col(
    transactionHeader.customerColumn,
  );
  const worksheetKeys = Object.keys(transactionSheet);
  let transactionCount = 0;
  for (const address of worksheetKeys) {
    if (!address.startsWith(customerColumnLetter)) continue;
    const rowNumber = Number(address.slice(customerColumnLetter.length));
    if (
      !Number.isInteger(rowNumber) ||
      rowNumber <= transactionHeader.headerRow + 1
    ) {
      continue;
    }
    const row = rowNumber - 1;
    const customer = String(
      readWorksheetValue(
        transactionSheet,
        row,
        transactionHeader.customerColumn,
      ) ?? "",
    ).trim();
    const category = normalizeText(
      String(
        readWorksheetValue(
          transactionSheet,
          row,
          transactionHeader.categoryColumn,
        ) ?? "",
      ),
    );
    if (!customer || !category) continue;
    const productIndex = PRODUCT_REQUIREMENTS.findIndex((requirement) =>
      category.includes(requirement.category),
    );
    if (productIndex < 0) continue;

    const quantity = parseQuantity(
      readWorksheetValue(
        transactionSheet,
        row,
        transactionHeader.quantityColumn,
      ),
      rowNumber,
    );
    const customerKey = normalizeText(customer);
    const customerQuantities = quantitiesByCustomer.get(customerKey) ??
      Array.from({ length: PRODUCT_REQUIREMENTS.length }, () => 0);
    customerQuantities[productIndex] += quantity;
    quantitiesByCustomer.set(customerKey, customerQuantities);
    transactionCount += 1;
  }

  const rows = clients.map(({ code, customer, seller, rowIndex }) => {
    const quantities =
      quantitiesByCustomer.get(normalizeText(customer)) ??
      Array.from({ length: PRODUCT_REQUIREMENTS.length }, () => 0);
    const savedResults = baseHeader.productStatusColumns.map((column) =>
      savedFormulaResult(readWorksheetValue(baseSheet, rowIndex, column)),
    );
    const products = PRODUCT_REQUIREMENTS.map((requirement, index) => {
      const usesSavedResult = savedResults[index] !== null;
      return {
        quantity: quantities[index],
        meetsRequirement:
          savedResults[index] ?? quantities[index] >= requirement.minimum,
        usesSavedResult,
      };
    });
    const missingProducts = products.flatMap((product, index) =>
      product.meetsRequirement ? [] : [PRODUCT_REQUIREMENTS[index].label],
    );
    return {
      code,
      customer,
      seller,
      products,
      meetsAll: missingProducts.length === 0,
      missingProducts,
      hasProductSales: quantities.some((quantity) => quantity > 0),
    };
  });

  return {
    name: fileName,
    rows,
    sellers: [...new Set(rows.map((row) => row.seller))].sort((a, b) =>
      a.localeCompare(b, "es"),
    ),
    transactionCount,
    formulaResultRows: rows.filter((row) =>
      row.products.every((product) => product.usesSavedResult),
    ).length,
  };
}

function OptionSelect({
  value,
  onChange,
  options,
  placeholder,
  ariaLabel,
}: {
  value: string;
  onChange: (value: string) => void;
  options: { value: string; label: string }[];
  placeholder?: string;
  ariaLabel: string;
}) {
  return (
    <div className="select-wrap">
      <select
        aria-label={ariaLabel}
        value={value}
        onChange={(event) => onChange(event.target.value)}
      >
        {placeholder && <option value="">{placeholder}</option>}
        {options.map((option) => (
          <option key={option.value} value={option.value}>
            {option.label}
          </option>
        ))}
      </select>
      <ChevronDown size={15} aria-hidden="true" />
    </div>
  );
}

function MultiSelectFilter({
  options,
  selectedValues,
  onChange,
  label,
  selectionSummary,
  allowSelectAll = true,
}: {
  options: string[];
  selectedValues: string[];
  onChange: (values: string[]) => void;
  label: string;
  selectionSummary?: (values: string[]) => string;
  allowSelectAll?: boolean;
}) {
  const [query, setQuery] = useState("");
  const matchingOptions = options.filter((option) =>
    normalizeText(option).includes(normalizeText(query)),
  );
  const allSelected = options.length > 0 && selectedValues.length === options.length;
  const buttonLabel = selectionSummary
    ? selectionSummary(selectedValues)
    : selectedValues.length === 0 || allSelected
      ? "Todos"
      : selectedValues.length === 1
        ? selectedValues[0]
        : `${selectedValues.length} seleccionados`;

  function toggleValue(value: string) {
    onChange(
      selectedValues.includes(value)
        ? selectedValues.filter((selected) => selected !== value)
        : [...selectedValues, value],
    );
  }

  return (
    <details className="multi-filter">
      <summary className="multi-filter-trigger" aria-label={label}>
        <span>{buttonLabel}</span>
        <ChevronDown size={14} aria-hidden="true" />
      </summary>
      <div className="multi-filter-menu">
        <div className="multi-filter-actions">
          <span>Seleccioná uno o varios</span>
          {allowSelectAll && (
            <button
              type="button"
              onClick={() => onChange(allSelected ? [] : options)}
            >
              {allSelected ? "Limpiar" : "Todos"}
            </button>
          )}
        </div>
        <div className="multi-filter-search">
          <Search size={15} aria-hidden="true" />
          <input
            type="search"
            value={query}
            onChange={(event) => setQuery(event.target.value)}
            placeholder={`Buscar ${label.replace(/^Filtrar por /, "").toLowerCase()}…`}
            aria-label={`Buscar en ${label.replace(/^Filtrar por /, "").toLowerCase()}`}
          />
          {query && (
            <button
              type="button"
              aria-label="Borrar búsqueda"
              onClick={() => setQuery("")}
            >
              <X size={14} />
            </button>
          )}
        </div>
        <div className="multi-filter-options">
          {matchingOptions.map((option) => (
            <label className="multi-filter-option" key={option}>
              <input
                type="checkbox"
                checked={selectedValues.includes(option)}
                onChange={() => toggleValue(option)}
              />
              <span>{option}</span>
            </label>
          ))}
          {matchingOptions.length === 0 && (
            <p className="multi-filter-empty">
              No encontramos opciones para “{query}”.
            </p>
          )}
        </div>
        {selectedValues.length > 0 && (
          <button
            className="multi-filter-done"
            onClick={(event) => {
              event.currentTarget.closest("details")?.removeAttribute("open");
            }}
            type="button"
          >
            Listo
          </button>
        )}
      </div>
    </details>
  );
}

export default function Home() {
  const inputRef = useRef<HTMLInputElement>(null);
  const comparisonInputRef = useRef<HTMLInputElement>(null);
  const topProductsInputRef = useRef<HTMLInputElement>(null);
  const consolidationInputRef = useRef<HTMLInputElement>(null);
  const [darkMode, setDarkMode] = useState(false);
  const [themeReady, setThemeReady] = useState(false);
  const [workbook, setWorkbook] = useState<WorkbookData | null>(null);
  const [comparisonSources, setComparisonSources] = useState<ComparisonSource[]>([]);
  const [comparisonMode, setComparisonMode] =
    useState<FileComparisonMode>("periods");
  const [activeComparisonSource, setActiveComparisonSource] = useState("base");
  const [comparisonError, setComparisonError] = useState("");
  const [topProductsWorkbook, setTopProductsWorkbook] =
    useState<ProductReport | null>(null);
  const [topProductsLoading, setTopProductsLoading] = useState(false);
  const [topProductsError, setTopProductsError] = useState("");
  const [topProductsDragging, setTopProductsDragging] = useState(false);
  const [consolidationWorkbook, setConsolidationWorkbook] =
    useState<ConsolidatedWorkbook | null>(null);
  const [consolidationLoading, setConsolidationLoading] = useState(false);
  const [consolidationExporting, setConsolidationExporting] = useState(false);
  const [consolidationDragging, setConsolidationDragging] = useState(false);
  const [consolidationProgress, setConsolidationProgress] = useState("");
  const [consolidationError, setConsolidationError] = useState("");
  const [productSeller, setProductSeller] = useState("Todos");
  const [productStatus, setProductStatus] =
    useState<ProductReportFilter>("all");
  const [productSearch, setProductSearch] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [successMessage, setSuccessMessage] = useState("");
  const [saveDialogOpen, setSaveDialogOpen] = useState(false);
  const [saveComparisonName, setSaveComparisonName] = useState("");
  const [dragging, setDragging] = useState(false);
  const [search, setSearch] = useState("");
  const [rowField, setRowField] = useState("");
  const [columnFields, setColumnFields] = useState<string[]>([]);
  const [metricFields, setMetricFields] = useState<string[]>([]);
  const metricField = metricFields[0] ?? "";
  const [dateGrouping, setDateGrouping] = useState<Grouping>("month");
  const [metricMode, setMetricMode] = useState<MetricMode>("sum");
  const [filters, setFilters] = useState<DimensionFilter[]>([]);
  const [savedComparisons, setSavedComparisons] = useState<SavedComparison[]>([]);
  const [activePage, setActivePage] = useState<
    "overview" | "data" | "top-products" | "consolidate"
  >("overview");
  const productSellerRows = useMemo(
    () =>
      topProductsWorkbook?.rows.filter(
        (row) => productSeller === "Todos" || row.seller === productSeller,
      ) ?? [],
    [productSeller, topProductsWorkbook],
  );
  const visibleProductRows = useMemo(() => {
    const query = normalizeText(productSearch);
    return productSellerRows.filter((row) => {
      const matchesSearch =
        !query ||
        normalizeText(`${row.code} ${row.customer} ${row.seller}`).includes(
          query,
        );
      const matchesStatus =
        productStatus === "all" ||
        (productStatus === "complete" && row.meetsAll) ||
        (productStatus === "incomplete" &&
          !row.meetsAll &&
          row.hasProductSales) ||
        (productStatus === "no-sales" && !row.hasProductSales);
      return matchesSearch && matchesStatus;
    });
  }, [productSearch, productSellerRows, productStatus]);
  const productReportStats = useMemo(
    () => ({
      total: productSellerRows.length,
      complete: productSellerRows.filter((row) => row.meetsAll).length,
      noSales: productSellerRows.filter((row) => !row.hasProductSales).length,
      incomplete: productSellerRows.filter(
        (row) => !row.meetsAll && row.hasProductSales,
      ).length,
    }),
    [productSellerRows],
  );
  const comparisonSourceField = useMemo(() => {
    if (!workbook) return "Archivo";
    let field = "Archivo";
    let suffix = 2;
    while (workbook.headers.includes(field)) {
      field = `Archivo ${suffix}`;
      suffix += 1;
    }
    return field;
  }, [workbook]);
  const analysisWorkbook = useMemo(() => {
    if (!workbook) return null;
    if (comparisonSources.length === 0) return workbook;
    const rows = [
      ...workbook.rows.map((row) => ({
        ...row,
        [comparisonSourceField]: "base",
      })),
      ...comparisonSources.flatMap((source) =>
        source.rows.map((sourceRow) => {
          const row: DataRow = { [comparisonSourceField]: source.id };
          workbook.headers.forEach((header) => {
            const sourceHeader = source.fieldMapping[header];
            row[header] = sourceHeader ? sourceRow[sourceHeader] ?? null : null;
          });
          return row;
        }),
      ),
    ];
    return {
      ...workbook,
      rows,
      headers: [...workbook.headers, comparisonSourceField],
    };
  }, [workbook, comparisonSources, comparisonSourceField]);

  useEffect(() => {
    const frame = window.requestAnimationFrame(() => {
      setDarkMode(window.localStorage.getItem("compara-theme") === "dark");
      try {
        const storedComparisons = window.localStorage.getItem(SAVED_COMPARISONS_KEY);
        if (storedComparisons) {
          const parsed: unknown = JSON.parse(storedComparisons);
          if (!Array.isArray(parsed) || !parsed.every(isSavedComparison)) {
            throw new Error("Las comparativas guardadas tienen un formato inválido.");
          }
          setSavedComparisons(parsed);
        }
      } catch (storageError) {
        setError(
          storageError instanceof Error
            ? storageError.message
            : "No pudimos recuperar las comparativas guardadas.",
        );
      }
      setThemeReady(true);
    });
    return () => window.cancelAnimationFrame(frame);
  }, []);

  useEffect(() => {
    if (!themeReady) return;
    window.localStorage.setItem("compara-theme", darkMode ? "dark" : "light");
  }, [darkMode, themeReady]);

  const dateField = useMemo(
    () =>
      analysisWorkbook
        ? findHeader(analysisWorkbook.headers, /fecha|date|mes|periodo|period/)
        : "",
    [analysisWorkbook],
  );
  const monthField = useMemo(
    () => (analysisWorkbook ? findHeader(analysisWorkbook.headers, /^mes$/) : ""),
    [analysisWorkbook],
  );
  const sellerField = useMemo(
    () =>
      analysisWorkbook
        ? findHeader(analysisWorkbook.headers, /vendedor|vendedora|seller|comercial|ejecutivo/)
        : "",
    [analysisWorkbook],
  );
  const productField = useMemo(
    () =>
      analysisWorkbook
        ? findHeader(analysisWorkbook.headers, /descripcion|producto|articulo|product|sku/)
        : "",
    [analysisWorkbook],
  );
  const comparisonMetricField = useMemo(
    () =>
      analysisWorkbook
        ? findHeader(
            analysisWorkbook.headers,
            /precio|importe|total|venta|factur|cantidad|peso/,
          )
        : "",
    [analysisWorkbook],
  );
  const numericFields = useMemo(() => {
    if (!analysisWorkbook) return [];
    return analysisWorkbook.headers.filter(
      (header) =>
        !/comprobante|codigo|cod\.?|cliente|documento|cuit|dni|id\b|postal/i.test(
          header,
        ) &&
        analysisWorkbook.rows.some((row) => typeof row[header] === "number"),
    );
  }, [analysisWorkbook]);
  const columnMetricFields = columnFields.filter((field) =>
    numericFields.includes(field),
  );
  const columnDimensionFields = columnFields.filter(
    (field) => !columnMetricFields.includes(field),
  );
  const activeMetricFields =
    columnMetricFields.length > 0 ? columnMetricFields : metricFields;

  const filterOptionsByField = useMemo(() => {
    if (!analysisWorkbook) return {};
    const valuesByField: Record<string, string[]> = {};
    filters.forEach(({ field }) => {
      if (!field || valuesByField[field]) return;
      const unique = new Map<string, string>();
      analysisWorkbook.rows.forEach((row) => {
        const value = displayGroupingValue(cellText(row[field] ?? null));
        const key = normalizeGroupingValue(value);
        if (!unique.has(key)) unique.set(key, value);
      });
      valuesByField[field] = [...unique.values()].sort((a, b) => {
        const monthA = MONTH_LOOKUP[normalizeText(a)];
        const monthB = MONTH_LOOKUP[normalizeText(b)];
        if (monthA !== undefined && monthB !== undefined) return monthA - monthB;
        return a.localeCompare(b, "es", { numeric: true, sensitivity: "base" });
      });
    });
    return valuesByField;
  }, [analysisWorkbook, filters]);

  const filteredRows = useMemo(() => {
    if (!analysisWorkbook) return [];
    const activeSourceId =
      comparisonMode === "separate" && comparisonSources.length
        ? activeComparisonSource
        : undefined;
    return analysisWorkbook.rows.filter((row) =>
      (!activeSourceId || row[comparisonSourceField] === activeSourceId) &&
      filters.every(({ field, values }) => {
        if (!field || values.length === 0) return true;
        const normalizedValue = normalizeGroupingValue(
          cellText(row[field] ?? null),
        );
        return values.some(
          (value) => normalizeGroupingValue(value) === normalizedValue,
        );
      }),
    );
  }, [
    analysisWorkbook,
    filters,
    comparisonMode,
    comparisonSources,
    activeComparisonSource,
    comparisonSourceField,
  ]);

  const pivotColumnFields =
    comparisonSources.length > 0 && comparisonMode === "periods"
      ? [...columnFields, comparisonSourceField]
      : columnFields;
  const pivotColumnDimensionFields = pivotColumnFields.filter(
    (field) => !numericFields.includes(field),
  );

  const pivot = useMemo(() => {
    if (!analysisWorkbook || !rowField || activeMetricFields.length === 0) return null;
    const rowTotals = new Map<string, PivotRowValues>();
    const rowOrder = new Set<string>();
    const columnOrder = new Set<string>();
    const totals = new Map<string, MetricValues>();
    const grandTotals: MetricValues = new Map();
    const rowLabelsByKey = new Map<string, string>();
    const normalizedColumnLabels = new Map<string, string>();
    const columnLabelsByKey = new Map<string, string>();
    const columnValuesByKey = new Map<string, string[]>();
    const getFieldValue = (row: DataRow, field: string) => {
      const value =
        field === comparisonSourceField && typeof row[field] === "string"
          ? row[field] === "base"
            ? workbook?.name.replace(/\.[^.]+$/, "") ?? "Archivo base"
            : comparisonSources.find((source) => source.id === row[field])?.label ??
              "Archivo"
          : row[field] ?? null;
      if (field === dateField) {
        const date = asDate(value);
        return date ? dateValue(date, dateGrouping) : "";
      }
      return cellText(value);
    };
    const getNormalizedLabel = (labels: Map<string, string>, value: string) => {
      const display = displayGroupingValue(value);
      const key = normalizeGroupingValue(display);
      if (!labels.has(key)) labels.set(key, display);
      return labels.get(key)!;
    };
    filteredRows.forEach((row) => {
      const rowValue = getFieldValue(row, rowField);
      if (!rowValue) return;
      const rowLabel =
        rowField === dateField
          ? rowValue
          : getNormalizedLabel(rowLabelsByKey, rowValue);
      const columnValues = pivotColumnDimensionFields.map((field) => {
        const value = getFieldValue(row, field);
        return field === dateField
          ? value
          : getNormalizedLabel(normalizedColumnLabels, value);
      });
      if (columnValues.some((value) => !value)) return;
      const columnKey = JSON.stringify(columnValues.map(normalizeGroupingValue));
      const columnLabel = pivotColumnDimensionFields.length
        ? pivotColumnDimensionFields
            .map((field, index) => `${field}: ${columnValues[index]}`)
            .join(" · ")
        : "Total";
      if (!columnLabelsByKey.has(columnKey)) {
        columnLabelsByKey.set(columnKey, columnLabel);
        columnValuesByKey.set(columnKey, columnValues);
      }
      const validMetrics = activeMetricFields.flatMap((field) => {
        const rawValue = row[field];
        const value =
          typeof rawValue === "number"
            ? rawValue
            : typeof rawValue === "string" && rawValue.trim()
              ? Number(rawValue)
              : Number.NaN;
        return Number.isFinite(value) ? [{ field, value }] : [];
      });
      if (validMetrics.length === 0) return;
      rowOrder.add(rowLabel);
      columnOrder.add(columnKey);

      if (!rowTotals.has(rowLabel)) rowTotals.set(rowLabel, new Map());
      const rowValues = rowTotals.get(rowLabel)!;
      if (!rowValues.has(columnKey)) rowValues.set(columnKey, new Map());
      const rowMetrics = rowValues.get(columnKey)!;
      if (!totals.has(columnKey)) totals.set(columnKey, new Map());
      const columnMetrics = totals.get(columnKey)!;
      validMetrics.forEach(({ field, value }) => {
        const existing = rowMetrics.get(field) ?? { total: 0, count: 0 };
        rowMetrics.set(field, {
          total: existing.total + value,
          count: existing.count + 1,
        });
        const columnTotal = columnMetrics.get(field) ?? { total: 0, count: 0 };
        columnMetrics.set(field, {
          total: columnTotal.total + value,
          count: columnTotal.count + 1,
        });
        const grandTotal = grandTotals.get(field) ?? { total: 0, count: 0 };
        grandTotals.set(field, {
          total: grandTotal.total + value,
          count: grandTotal.count + 1,
        });
      });
    });

    const rowLabels = [...rowOrder].sort((a, b) => {
      if (rowField === dateField) return dateSort(a) - dateSort(b);
      const monthA = MONTH_LOOKUP[normalizeText(a)];
      const monthB = MONTH_LOOKUP[normalizeText(b)];
      if (monthA !== undefined && monthB !== undefined) return monthA - monthB;
      return a.localeCompare(b, "es", { numeric: true, sensitivity: "base" });
    });
    const compareLabels = (a: string, b: string) => {
      const monthA = MONTH_LOOKUP[normalizeText(a)];
      const monthB = MONTH_LOOKUP[normalizeText(b)];
      if (monthA !== undefined && monthB !== undefined) return monthA - monthB;
      const dateA = dateSort(a);
      const dateB = dateSort(b);
      if (dateA !== Number.MAX_SAFE_INTEGER && dateB !== Number.MAX_SAFE_INTEGER) {
        return dateA - dateB;
      }
      return a.localeCompare(b, "es", { numeric: true, sensitivity: "base" });
    };
    const columnLabels = [...columnOrder].sort((a, b) => {
      const valuesA = columnValuesByKey.get(a) ?? [];
      const valuesB = columnValuesByKey.get(b) ?? [];
      for (let index = 0; index < Math.min(valuesA.length, valuesB.length); index += 1) {
        const comparison = compareLabels(valuesA[index], valuesB[index]);
        if (comparison !== 0) return comparison;
      }
      return valuesA.length - valuesB.length;
    });
    return {
      rowLabels,
      columnLabels,
      columnDisplayLabels: columnLabelsByKey,
      rowTotals,
      totals,
      grandTotals,
    };
  }, [
    analysisWorkbook,
    filteredRows,
    rowField,
    pivotColumnDimensionFields,
    activeMetricFields,
    dateField,
    dateGrouping,
    workbook,
    comparisonSources,
    comparisonSourceField,
  ]);

  const comparisonPeriods = useMemo(() => {
    if (comparisonSources.length > 0 && comparisonMode === "periods") return [];
    if (
      !monthField ||
      columnDimensionFields.length !== 1 ||
      columnDimensionFields[0] !== monthField ||
      filters.filter((filter) => filter.field === monthField).length !== 1
    ) {
      return [];
    }
    const values = filters.find((filter) => filter.field === monthField)?.values ?? [];
    if (values.length !== 2) return [];
    return [...values].sort(
      (a, b) =>
        (MONTH_LOOKUP[normalizeText(a)] ?? Number.MAX_SAFE_INTEGER) -
        (MONTH_LOOKUP[normalizeText(b)] ?? Number.MAX_SAFE_INTEGER),
    ).map((value) => JSON.stringify([normalizeGroupingValue(value)]));
  }, [
    monthField,
    columnDimensionFields,
    filters,
    comparisonSources.length,
    comparisonMode,
  ]);

  const shownRows = useMemo(() => {
    if (!analysisWorkbook) return [];
    const query = normalizeText(search);
    return analysisWorkbook.rows.filter((row) =>
      !query ||
      analysisWorkbook.headers.some((header) =>
        normalizeText(cellText(row[header] ?? null)).includes(query),
      ),
    );
  }, [analysisWorkbook, search]);

  const quickMetrics = useMemo(() => {
    if (!analysisWorkbook) return { sellers: 0, products: 0, customers: 0 };
    const uniqueCount = (pattern: RegExp) => {
      const field = findHeader(analysisWorkbook.headers, pattern);
      return field
        ? new Set(
            analysisWorkbook.rows
              .map((row) => cellText(row[field] ?? null))
              .filter((value) => value !== "Sin dato")
              .map(normalizeGroupingValue),
          )
        .size
        : 0;
    };
    return {
      sellers: uniqueCount(/vendedor|vendedora|seller|comercial|ejecutivo/),
      products: uniqueCount(/articulo|producto|descripcion|description|sku/),
      customers: uniqueCount(/cliente|customer/),
    };
  }, [analysisWorkbook]);

  const dateWarnings = useMemo(() => {
    if (!analysisWorkbook || !dateField) return 0;
    if (!monthField || monthField === dateField) return 0;
    return analysisWorkbook.rows.reduce((count, row) => {
      const date = asDate(row[dateField] ?? null);
      const month = MONTH_LOOKUP[normalizeText(cellText(row[monthField] ?? null))];
      return date && month !== undefined && date.getMonth() !== month
        ? count + 1
        : count;
    }, 0);
  }, [analysisWorkbook, dateField, monthField]);

  const chartItems = useMemo(() => {
    if (!pivot) return [];
    return pivot.rowLabels.map((label) => {
      const cells = pivot.columnLabels
        .map((column) =>
          pivot.rowTotals.get(label)?.get(column)?.get(metricField),
        )
        .filter((cell): cell is PivotCell => Boolean(cell));
      const total = cells.reduce((sum, cell) => sum + cell.total, 0);
      const count = cells.reduce((sum, cell) => sum + cell.count, 0);
      const value =
        metricMode === "count"
          ? count
          : metricMode === "average"
            ? total / Math.max(1, count)
            : total;
      return { label, value };
    }).slice(-12);
  }, [pivot, metricField, metricMode]);

  const maximumChartValue = Math.max(1, ...chartItems.map((item) => Math.abs(item.value)));
  const formatMetricValue = (field: string, value: number) =>
    /precio|importe|total|venta|factur/i.test(field)
      ? formatMoney(value)
      : formatNumber(value, 2);
  const formatPivotMetric = (field: string, value: number) =>
    metricMode === "count"
      ? formatNumber(value)
      : formatMetricValue(field, value);

  async function loadFile(file: File) {
    setLoading(true);
    setError("");
    setSuccessMessage("");
    try {
      if (file.size > MAX_UPLOAD_BYTES) {
        throw new Error("El archivo supera el límite de 50 MB.");
      }
      const data = await file.arrayBuffer();
      const parsedWorkbook = XLSX.read(data, { type: "array", cellDates: true });
      const sheets = parsedWorkbook.SheetNames;
      const sheet = sheets[0];
      if (!sheet || !parsedWorkbook.Sheets[sheet]) {
        throw new Error("No encontramos hojas con datos en este archivo.");
      }
      const parsed = parseWorkbookSheet(parsedWorkbook, sheet);
      if (!parsed.headers.length || !parsed.rows.length) {
        throw new Error("La primera hoja no tiene filas de datos debajo de los encabezados.");
      }
      const nextWorkbook: WorkbookData = {
        name: file.name,
        sheet,
        sheets,
        buffer: data,
        ...parsed,
      };
      const detectedDate = findHeader(parsed.headers, /fecha|date|mes|periodo|period/);
      const detectedMetric = findHeader(
        parsed.headers,
        /precio|importe|total|venta|factur|cantidad|peso/,
      );
      setWorkbook(nextWorkbook);
      setComparisonSources([]);
      setComparisonMode("periods");
      setActiveComparisonSource("base");
      setComparisonError("");
      setRowField(
        findHeader(parsed.headers, /vendedor|vendedora|seller|comercial|ejecutivo/) ||
          detectedDate ||
          parsed.headers[0],
      );
      setColumnFields([findHeader(parsed.headers, /^mes$/)].filter(Boolean));
      setMetricFields(
        parsed.headers.includes(detectedMetric) && detectedMetric ? [detectedMetric] : [],
      );
      setMetricMode("sum");
      setDateGrouping("month");
      const detectedMonth = findHeader(parsed.headers, /^mes$/);
      setFilters(detectedMonth ? [{ field: detectedMonth, values: [] }] : []);
      setActivePage("overview");
    } catch (loadError) {
      const message =
        loadError instanceof Error
          ? loadError.message
          : "No pudimos leer el archivo. Probá exportarlo nuevamente como Excel.";
      setError(message);
    } finally {
      setLoading(false);
    }
  }

  function handleFileChange(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (file) void loadFile(file);
    event.target.value = "";
  }

  async function handleComparisonFiles(event: ChangeEvent<HTMLInputElement>) {
    const files = Array.from(event.target.files ?? []);
    event.target.value = "";
    if (!files.length || !workbook) return;
    setComparisonError("");
    const loadedSources: ComparisonSource[] = [];
    try {
      for (const file of files) {
        if (file.size > MAX_UPLOAD_BYTES) {
          throw new Error(`${file.name}: el archivo supera el límite de 50 MB.`);
        }
        const buffer = await file.arrayBuffer();
        const parsedWorkbook = XLSX.read(buffer, {
          type: "array",
          cellDates: true,
        });
        const sheet = parsedWorkbook.SheetNames[0];
        if (!sheet || !parsedWorkbook.Sheets[sheet]) {
          throw new Error(`${file.name}: no encontramos hojas con datos.`);
        }
        const parsed = parseWorkbookSheet(parsedWorkbook, sheet);
        if (!parsed.headers.length || !parsed.rows.length) {
          throw new Error(`${file.name}: la primera hoja no tiene filas de datos.`);
        }
        const fieldMapping = Object.fromEntries(
          workbook.headers.map((canonicalHeader) => {
            const matchingHeader = parsed.headers.find(
              (header) => normalizeText(header) === normalizeText(canonicalHeader),
            );
            return [canonicalHeader, matchingHeader ?? ""];
          }),
        );
        loadedSources.push({
          id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
          name: file.name,
          label: file.name.replace(/\.[^.]+$/, ""),
          sheet,
          sheets: parsedWorkbook.SheetNames,
          buffer,
          ...parsed,
          fieldMapping,
        });
      }
      setComparisonSources((current) => [...current, ...loadedSources]);
      setComparisonMode("periods");
      setActiveComparisonSource("base");
    } catch (fileError) {
      setComparisonError(
        fileError instanceof Error
          ? fileError.message
          : "No pudimos leer los archivos seleccionados.",
      );
    }
  }

  function updateSourceMapping(
    sourceId: string,
    targetField: string,
    sourceField: string,
  ) {
    setComparisonSources((current) =>
      current.map((source) =>
        source.id === sourceId
          ? {
              ...source,
              fieldMapping: {
                ...source.fieldMapping,
                [targetField]: sourceField,
              },
            }
          : source,
      ),
    );
  }

  function updateSourceLabel(sourceId: string, label: string) {
    setComparisonSources((current) =>
      current.map((source) =>
        source.id === sourceId ? { ...source, label } : source,
      ),
    );
  }

  function removeComparisonSource(sourceId: string) {
    const remaining = comparisonSources.filter((source) => source.id !== sourceId);
    setComparisonSources(remaining);
    if (!remaining.length) setComparisonMode("periods");
    if (activeComparisonSource === sourceId) setActiveComparisonSource("base");
  }

  async function loadTopProductsFile(file: File) {
    setTopProductsLoading(true);
    setTopProductsError("");
    try {
      if (file.size > MAX_UPLOAD_BYTES) {
        throw new Error("El archivo supera el límite de 50 MB.");
      }
      const buffer = await file.arrayBuffer();
      const parsedWorkbook = XLSX.read(buffer, {
        type: "array",
        cellDates: true,
        cellFormula: true,
      });
      setTopProductsWorkbook(createProductReport(parsedWorkbook, file.name));
      setProductSeller("Todos");
      setProductStatus("all");
      setProductSearch("");
    } catch (fileError) {
      setTopProductsError(
        fileError instanceof Error
          ? fileError.message
          : "No pudimos leer el archivo. Probá exportarlo nuevamente como Excel.",
      );
    } finally {
      setTopProductsLoading(false);
    }
  }

  function handleTopProductsFileChange(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (file) void loadTopProductsFile(file);
    event.target.value = "";
  }

  function handleTopProductsDrop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault();
    setTopProductsDragging(false);
    const file = event.dataTransfer.files?.[0];
    if (file) void loadTopProductsFile(file);
  }

  function downloadTopProductsReport() {
    if (!topProductsWorkbook) return;
    const headers = [
      "Código",
      "Cliente",
      "Vendedor",
      ...PRODUCT_REQUIREMENTS.flatMap((product) => [
        `${product.label} estado`,
        `${product.label} unidades calculadas`,
      ]),
      "Productos faltantes",
      "Cumple los 5",
    ];
    const detailRows = visibleProductRows.map((row) => [
      row.code,
      row.customer,
      row.seller,
      ...row.products.flatMap((product) => [
        product.meetsRequirement ? "Cumple" : "No cumple",
        product.usesSavedResult ? "" : product.quantity,
      ]),
      row.missingProducts.join(", ") || "Ninguno",
      row.meetsAll ? "Sí" : "No",
    ]);
    const summarySheet = XLSX.utils.aoa_to_sheet([
      ["Archivo", topProductsWorkbook.name],
      ["Vendedor", productSeller],
      ["Clientes incluidos", visibleProductRows.length],
      ["Cumplen los 5 productos", visibleProductRows.filter((row) => row.meetsAll).length],
      ["Vendieron, pero les faltan productos", visibleProductRows.filter((row) => !row.meetsAll && row.hasProductSales).length],
      ["No vendieron ninguno de los 5", visibleProductRows.filter((row) => !row.hasProductSales).length],
      ["Productos vendidos", "Café ≥ 6 unidades; Edulco, Azúcar, Yerba y Té ≥ 1 unidad"],
    ]);
    const detailSheet = XLSX.utils.aoa_to_sheet([headers, ...detailRows]);
    const report = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(report, summarySheet, "Resumen");
    XLSX.utils.book_append_sheet(report, detailSheet, "Detalle");
    XLSX.writeFile(report, "informe-cumplimiento-5-productos.xlsx");
  }

  function handleDrop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault();
    setDragging(false);
    const file = event.dataTransfer.files?.[0];
    if (file) void loadFile(file);
  }

  function handleSheetChange(sheet: string) {
    if (!workbook) return;
    const parsedWorkbook = XLSX.read(workbook.buffer, {
      type: "array",
      cellDates: true,
    });
    const parsed = parseWorkbookSheet(parsedWorkbook, sheet);
    setWorkbook({ ...workbook, sheet, ...parsed });
    setComparisonSources([]);
    setComparisonMode("periods");
    setActiveComparisonSource("base");
    setComparisonError("");
    const detectedDate = findHeader(parsed.headers, /fecha|date|mes|periodo|period/);
    const detectedMetric = findHeader(
      parsed.headers,
      /precio|importe|total|venta|factur|cantidad|peso/,
    );
    setRowField(
      findHeader(parsed.headers, /vendedor|vendedora|seller|comercial|ejecutivo/) ||
        detectedDate ||
        parsed.headers[0] ||
        "",
    );
    setColumnFields([findHeader(parsed.headers, /^mes$/)].filter(Boolean));
    setMetricFields(
      parsed.headers.includes(detectedMetric) && detectedMetric ? [detectedMetric] : [],
    );
    const detectedMonth = findHeader(parsed.headers, /^mes$/);
    setFilters(detectedMonth ? [{ field: detectedMonth, values: [] }] : []);
  }

  function exportPivot() {
    if (!pivot || !workbook) return;
    const rows = pivot.rowLabels.map((label) => {
      const row: Record<string, string | number> = { [rowField]: label };
      pivot.columnLabels.forEach((column) => {
        const metricValues = pivot.rowTotals.get(label)?.get(column);
        activeMetricFields.forEach((field) => {
          const cell = metricValues?.get(field);
          const header = pivotColumnDimensionFields.length
            ? `${pivot.columnDisplayLabels.get(column) ?? column} · ${field}`
            : field;
          row[header] = pivotCellValue(cell, metricMode);
        });
      });
      activeMetricFields.forEach((field) => {
        if (comparisonPeriods.length === 2) {
          const values = comparisonPeriods.map((period) =>
            pivotCellValue(
              pivot.rowTotals.get(label)?.get(period)?.get(field),
              metricMode,
            ),
          );
          const difference = values[1] - values[0];
          row[`${field} · Variación`] = difference;
          row[`${field} · Cambio %`] =
            values[0] === 0
              ? "N/D"
              : `${formatNumber((difference / Math.abs(values[0])) * 100, 1)}%`;
        }
        const cells = pivot.columnLabels
          .map((column) => pivot.rowTotals.get(label)?.get(column)?.get(field))
          .filter((cell): cell is PivotCell => Boolean(cell));
        const total = cells.reduce((sum, cell) => sum + cell.total, 0);
        const count = cells.reduce((sum, cell) => sum + cell.count, 0);
        row[`${field} · Total`] =
          metricMode === "count"
            ? count
            : metricMode === "average"
              ? total / Math.max(1, count)
              : total;
      });
      return row;
    });
    const sheet = XLSX.utils.json_to_sheet(rows);
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, sheet, "Tabla dinámica");
    XLSX.writeFile(book, `comparativa-${workbook.name.replace(/\.[^.]+$/, "")}.xlsx`);
  }

  async function loadConsolidationFile(file: File) {
    setConsolidationLoading(true);
    setConsolidationWorkbook(null);
    setConsolidationError("");
    setConsolidationProgress("Preparando el archivo…");
    try {
      if (file.size > MAX_UPLOAD_BYTES) {
        throw new Error("El archivo supera el límite de 50 MB.");
      }
      await new Promise<void>((resolve) => window.setTimeout(resolve, 0));
      const buffer = await file.arrayBuffer();
      const sourceWorkbook = XLSX.read(buffer, { type: "array", cellDates: true });
      const schemas: {
        name: string;
        worksheet: XLSX.WorkSheet;
        range: XLSX.Range;
        headers: string[];
        outputIndexes: number[];
      }[] = [];
      const outputHeaders: string[] = [];
      const outputHeaderIndexes = new Map<string, number>();

      for (const name of sourceWorkbook.SheetNames) {
        const worksheet = sourceWorkbook.Sheets[name];
        if (!worksheet?.["!ref"]) continue;
        const range = XLSX.utils.decode_range(worksheet["!ref"]);
        if (range.e.r <= range.s.r) continue;
        const headerRows = XLSX.utils.sheet_to_json<unknown[]>(worksheet, {
          header: 1,
          range: {
            s: { r: range.s.r, c: range.s.c },
            e: { r: range.s.r, c: range.e.c },
          },
          raw: true,
          blankrows: false,
        });
        const rawHeaders = headerRows[0] ?? [];
        if (!rawHeaders.some((value) => String(value ?? "").trim())) continue;

        const seenHeaders = new Map<string, number>();
        const headers = rawHeaders.map((value, index) => {
          const header = cleanHeader(value, index);
          const key = normalizeText(header);
          const count = (seenHeaders.get(key) ?? 0) + 1;
          seenHeaders.set(key, count);
          return count > 1 ? `${header} (${count})` : header;
        });
        const outputIndexes = headers.map((header) => {
          const key = normalizeText(header);
          const existingIndex = outputHeaderIndexes.get(key);
          if (existingIndex !== undefined) return existingIndex;
          const index = outputHeaders.length;
          outputHeaders.push(header);
          outputHeaderIndexes.set(key, index);
          return index;
        });
        schemas.push({ name, worksheet, range, headers, outputIndexes });
      }

      if (schemas.length === 0) {
        throw new Error("No encontramos hojas con encabezados y filas de datos.");
      }
      let sourceHeader = "Hoja de origen";
      let sourceHeaderSuffix = 2;
      while (outputHeaderIndexes.has(normalizeText(sourceHeader))) {
        sourceHeader = `Hoja de origen ${sourceHeaderSuffix}`;
        sourceHeaderSuffix += 1;
      }
      if (outputHeaders.length + 1 > MAX_EXCEL_COLUMNS) {
        throw new Error("La unión supera el máximo de 16.384 columnas que admite Excel.");
      }

      const rows: Cell[][] = [];
      const includedSheets: string[] = [];
      for (const [sheetIndex, schema] of schemas.entries()) {
        const firstDataRow = schema.range.s.r + 1;
        let includedSheet = false;
        for (
          let startRow = firstDataRow;
          startRow <= schema.range.e.r;
          startRow += CONSOLIDATION_CHUNK_SIZE
        ) {
          const endRow = Math.min(
            startRow + CONSOLIDATION_CHUNK_SIZE - 1,
            schema.range.e.r,
          );
          setConsolidationProgress(
            `Leyendo hoja ${sheetIndex + 1} de ${schemas.length}: ${schema.name}`,
          );
          await new Promise<void>((resolve) => window.setTimeout(resolve, 0));
          const rawRows = XLSX.utils.sheet_to_json<unknown[]>(schema.worksheet, {
            header: 1,
            range: {
              s: { r: startRow, c: schema.range.s.c },
              e: { r: endRow, c: schema.range.e.c },
            },
            defval: null,
            raw: true,
            blankrows: false,
          });
          for (const rawRow of rawRows) {
            if (
              !rawRow.some(
                (value) =>
                  value !== null &&
                  value !== undefined &&
                  (typeof value !== "string" || value.trim() !== ""),
              )
            ) {
              continue;
            }
            if (rows.length >= MAX_EXCEL_ROWS - 1) {
              throw new Error(
                "La combinación supera el máximo de 1.048.576 filas por hoja de Excel. Dividí los datos en varios archivos para conservarlos todos.",
              );
            }
            const row: Cell[] = Array(outputHeaders.length + 1).fill(null);
            row[0] = schema.name;
            schema.outputIndexes.forEach((outputIndex, columnIndex) => {
              const value = rawRow[columnIndex];
              row[outputIndex + 1] =
                value instanceof Date ||
                typeof value === "string" ||
                typeof value === "number" ||
                typeof value === "boolean"
                  ? value
                  : null;
            });
            rows.push(row);
            includedSheet = true;
          }
        }
        if (includedSheet) includedSheets.push(schema.name);
      }

      if (rows.length === 0) {
        throw new Error("Las hojas no contienen filas con datos para consolidar.");
      }
      setConsolidationWorkbook({
        name: file.name,
        headers: [sourceHeader, ...outputHeaders],
        rows,
        sheets: includedSheets,
      });
      setConsolidationProgress("");
    } catch (fileError) {
      setConsolidationError(
        fileError instanceof Error
          ? fileError.message
          : "No pudimos consolidar el archivo. Verificá que sea un Excel válido.",
      );
      setConsolidationProgress("");
    } finally {
      setConsolidationLoading(false);
    }
  }

  function handleConsolidationFileChange(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (file) void loadConsolidationFile(file);
    event.target.value = "";
  }

  function handleConsolidationDrop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault();
    setConsolidationDragging(false);
    const file = event.dataTransfer.files?.[0];
    if (file) void loadConsolidationFile(file);
  }

  async function exportConsolidatedWorkbook() {
    if (!consolidationWorkbook) return;
    setConsolidationExporting(true);
    setConsolidationError("");
    try {
      const worksheet = XLSX.utils.aoa_to_sheet([consolidationWorkbook.headers]);
      for (
        let start = 0;
        start < consolidationWorkbook.rows.length;
        start += CONSOLIDATION_CHUNK_SIZE
      ) {
        const end = Math.min(
          start + CONSOLIDATION_CHUNK_SIZE,
          consolidationWorkbook.rows.length,
        );
        setConsolidationProgress(
          `Preparando descarga: ${end.toLocaleString("es-AR")} de ${consolidationWorkbook.rows.length.toLocaleString("es-AR")} filas`,
        );
        await new Promise<void>((resolve) => window.setTimeout(resolve, 0));
        XLSX.utils.sheet_add_aoa(
          worksheet,
          consolidationWorkbook.rows.slice(start, end),
          { origin: -1 },
        );
      }
      const outputWorkbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(outputWorkbook, worksheet, "Datos unificados");
      const baseName = consolidationWorkbook.name
        .replace(/\.[^.]+$/, "")
        .replace(/[<>:"|?*]/g, "-");
      XLSX.writeFile(outputWorkbook, `${baseName}-hojas-unidas.xlsx`);
      setConsolidationProgress("");
    } catch (exportError) {
      setConsolidationError(
        exportError instanceof Error
          ? `No pudimos crear el Excel unificado: ${exportError.message}`
          : "No pudimos crear el Excel unificado.",
      );
      setConsolidationProgress("");
    } finally {
      setConsolidationExporting(false);
    }
  }

  function persistSavedComparisons(comparisons: SavedComparison[]) {
    try {
      window.localStorage.setItem(
        SAVED_COMPARISONS_KEY,
        JSON.stringify(comparisons),
      );
      setSavedComparisons(comparisons);
      setError("");
      return true;
    } catch (storageError) {
      setError(
        storageError instanceof Error
          ? `No pudimos guardar la comparativa: ${storageError.message}`
          : "No pudimos guardar la comparativa en este navegador.",
      );
      return false;
    }
  }

  function openSaveComparisonDialog() {
    if (!workbook) return;
    const suggestedName = `${rowField || "Reporte"} · ${
      activeMetricFields.join(", ") || "Ventas"
    }`;
    setSaveComparisonName(suggestedName);
    setSaveDialogOpen(true);
  }

  function saveCurrentComparison(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const name = saveComparisonName.trim();
    if (!name) return;
    const comparison: SavedComparison = {
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      name,
      rowField,
      columnFields: [...columnFields],
      metricFields: [...metricFields],
      dateGrouping,
      metricMode,
      filters: filters.map((filter) => ({
        field: filter.field,
        values: [...filter.values],
      })),
    };
    if (persistSavedComparisons([...savedComparisons, comparison])) {
      setSaveDialogOpen(false);
      setSuccessMessage(`Comparativa “${name}” guardada.`);
    }
  }

  function applySavedComparison(comparison: SavedComparison) {
    if (!workbook) return;
    const availableFields = new Set(workbook.headers);
    const nextRowField = availableFields.has(comparison.rowField)
      ? comparison.rowField
      : workbook.headers[0] ?? "";
    const nextColumnFields = comparison.columnFields.filter(
      (field) => availableFields.has(field) && field !== nextRowField,
    );
    const nextMetricFields = comparison.metricFields.filter(
      (field) => numericFields.includes(field),
    );
    const nextFilters = comparison.filters.filter(({ field }) =>
      availableFields.has(field),
    );
    const missingFields = [
      ...(!availableFields.has(comparison.rowField) ? [comparison.rowField] : []),
      ...comparison.columnFields.filter((field) => !availableFields.has(field)),
      ...comparison.metricFields.filter(
        (field) => !numericFields.includes(field),
      ),
      ...comparison.filters
        .filter(({ field }) => !availableFields.has(field))
        .map(({ field }) => field),
    ];
    const unavailableFilterValues = nextFilters.flatMap(({ field, values }) => {
      const availableValues = new Set(
        workbook.rows.map((row) =>
          normalizeGroupingValue(cellText(row[field] ?? null)),
        ),
      );
      return values
        .filter((value) => !availableValues.has(normalizeGroupingValue(value)))
        .map((value) => `${field}: ${value}`);
    });
    setRowField(nextRowField);
    setColumnFields(nextColumnFields);
    setMetricFields(nextMetricFields);
    setDateGrouping(comparison.dateGrouping);
    setMetricMode(comparison.metricMode);
    setFilters(nextFilters);
    if (missingFields.length || unavailableFilterValues.length) {
      const warnings = [
        missingFields.length
          ? `No se encontraron estas columnas: ${[...new Set(missingFields)].join(", ")}.`
          : "",
        unavailableFilterValues.length
          ? `Algunos valores guardados no aparecen en este archivo: ${unavailableFilterValues.join(", ")}.`
          : "",
      ].filter(Boolean);
      setError(warnings.join(" "));
    } else {
      setError("");
    }
  }

  function removeSavedComparison(id: string) {
    const comparison = savedComparisons.find((item) => item.id === id);
    if (
      persistSavedComparisons(
        savedComparisons.filter((item) => item.id !== id),
      ) &&
      comparison
    ) {
      setSuccessMessage(`Comparativa “${comparison.name}” eliminada.`);
    }
  }

  function updateDimensionFilter(index: number, filter: DimensionFilter) {
    setFilters((current) =>
      current.map((existing, currentIndex) =>
        currentIndex === index ? filter : existing,
      ),
    );
  }

  function addDimensionFilter(field = "") {
    setFilters((current) => {
      if (field && current.some((filter) => filter.field === field)) {
        return current;
      }
      const nextField =
        field ||
        workbook?.headers.find(
          (header) => !current.some((filter) => filter.field === header),
        ) ||
        "";
      return nextField ? [...current, { field: nextField, values: [] }] : current;
    });
  }

  function removeDimensionFilter(index: number) {
    setFilters((current) => current.filter((_, currentIndex) => currentIndex !== index));
  }

  function updateMetricFields(fields: string[]) {
    setMetricFields(fields);
  }

  function updateColumnFields(fields: string[]) {
    const nextFields = fields.filter((field) => field !== rowField);
    setColumnFields(nextFields);
  }

  function updateRowField(field: string) {
    setRowField(field);
    setColumnFields((current) => current.filter((column) => column !== field));
  }

  function applyComparisonPreset(kind: "sellers" | "products" | "evolution") {
    const nextRowField =
      kind === "sellers"
        ? sellerField
        : kind === "products"
          ? productField
          : dateField;
    setRowField(nextRowField || workbook?.headers[0] || "");
    setColumnFields(kind === "evolution" ? [] : [monthField || dateField].filter(Boolean));
    setDateGrouping("month");
    setMetricFields(
      metricFields.length > 0
        ? metricFields
        : comparisonMetricField
          ? [comparisonMetricField]
          : [],
    );
    setMetricMode("sum");
    setFilters(monthField ? [{ field: monthField, values: [] }] : []);
  }

  const selectorOptions = (fields: string[]) =>
    fields.map((field) => ({ value: field, label: field }));
  const activeFilterSummary = filters.flatMap((filter) => {
    const options = filterOptionsByField[filter.field] ?? [];
    if (!filter.values.length || filter.values.length === options.length) return [];
    return [`${filter.field}: ${filter.values.join(", ")}`];
  });
  const valueColumns =
    pivot?.columnLabels.flatMap((period) =>
      activeMetricFields.map((field) => ({
        key: `${period}::${field}`,
        period,
        field,
      })),
    ) ?? [];

  function metricTotal(cells: (PivotCell | undefined)[]) {
    return cells.reduce<PivotCell>(
      (result, cell) => ({
        total: result.total + (cell?.total ?? 0),
        count: result.count + (cell?.count ?? 0),
      }),
      { total: 0, count: 0 },
    );
  }

  return (
    <main className={`app-shell ${darkMode ? "theme-dark" : ""}`}>
      <aside className="sidebar">
        <a className="brand" href="#" aria-label="Compara, inicio">
          <span className="brand-mark"><Coffee size={19} strokeWidth={2.2} /></span>
          <span className="brand-name">compara<span>.</span></span>
        </a>

        <div className="workspace-label">ESPACIO DE TRABAJO</div>
        <button className="workspace-switcher" type="button">
          <span className="workspace-avatar">C</span>
          <span className="workspace-copy"><strong>Café Santos</strong><small>Mi espacio</small></span>
          <ChevronDown size={15} />
        </button>

        <div className="nav-label">ANÁLISIS</div>
        <nav className="side-nav" aria-label="Navegación principal">
          <button
            className={`nav-item ${activePage === "overview" ? "active" : ""}`}
            onClick={() => setActivePage("overview")}
            type="button"
          >
            <BarChart3 size={18} /><span>Vista general</span>
          </button>
          <button
            className={`nav-item ${activePage === "data" ? "active" : ""}`}
            onClick={() => setActivePage("data")}
            type="button"
          >
            <Table2 size={18} /><span>Explorar datos</span>
          </button>
          <button
            className={`nav-item ${activePage === "top-products" ? "active" : ""}`}
            onClick={() => setActivePage("top-products")}
            type="button"
          >
            <Award size={18} /><span>Los 5 productos</span>
          </button>
          <button
            className={`nav-item ${activePage === "consolidate" ? "active" : ""}`}
            onClick={() => setActivePage("consolidate")}
            type="button"
          >
            <Layers3 size={18} /><span>Unir hojas</span>
          </button>
        </nav>

        <div className="sidebar-divider" />
        <div className="nav-label">TUS ARCHIVOS</div>
        {workbook ? (
          <button className="file-nav-item" type="button" title={workbook.name}>
            <FileSpreadsheet size={16} />
            <span>{workbook.name}</span>
            <span className="file-dot" />
          </button>
        ) : (
          <div className="no-files">Tus archivos aparecerán acá.</div>
        )}
        {comparisonSources.map((source) => (
          <button
            className="file-nav-item"
            key={source.id}
            type="button"
            title={source.name}
          >
            <FileSpreadsheet size={16} />
            <span>{source.name}</span>
            <span className="file-dot" />
          </button>
        ))}
        <button className="add-file-button" onClick={() => inputRef.current?.click()} type="button">
          <Plus size={15} /> Agregar archivo
        </button>

        <div className="sidebar-bottom">
          <div className="privacy-card">
            <span className="privacy-icon"><LockKeyhole size={15} /></span>
            <div><strong>Tus datos, tuyos.</strong><p>Los archivos se procesan en tu navegador.</p></div>
          </div>
          <button className="profile-button" type="button">
            <span className="profile-avatar">RP</span>
            <span className="profile-copy"><strong>Graciela</strong><small>Plan inicial</small></span>
            <Settings2 size={16} />
          </button>
        </div>
      </aside>

      <section className="main-panel">
        <header className="topbar">
          <div className="breadcrumb"><span>Mi espacio</span><span className="crumb-separator">/</span><strong>{activePage === "overview" ? "Vista general" : activePage === "data" ? "Explorar datos" : activePage === "top-products" ? "Los 5 productos" : "Unir hojas"}</strong></div>
          <div className="topbar-actions">
            <button
              className="theme-toggle"
              onClick={() => setDarkMode((enabled) => !enabled)}
              type="button"
              aria-label={darkMode ? "Activar modo claro" : "Activar modo oscuro"}
              aria-pressed={darkMode}
              title={darkMode ? "Activar modo claro" : "Activar modo oscuro"}
            >
              {darkMode ? <Sun size={17} /> : <Moon size={17} />}
              <span>{darkMode ? "Modo claro" : "Modo oscuro"}</span>
            </button>
            <span className="secure-badge"><span /> Procesamiento local</span>
            <button className="help-button" type="button" aria-label="Ayuda"><CircleHelp size={18} /></button>
            <button className="top-avatar" type="button" aria-label="Perfil">RP</button>
          </div>
        </header>

        <div className="content">
          {activePage === "consolidate" ? (
            <section className="consolidate-screen">
              <div className="welcome-topline"><span className="eyebrow-dot" /> HERRAMIENTA PARA TUS EXCEL</div>
              <h1>Uní todas tus hojas.</h1>
              <p className="welcome-description">
                Cargá un Excel y reuní las filas de sus hojas en una sola planilla,
                lista para analizar.
              </p>
              <div
                className={`upload-card ${consolidationDragging ? "dragging" : ""} ${consolidationLoading ? "is-loading" : ""}`}
                onDragEnter={(event) => { event.preventDefault(); setConsolidationDragging(true); }}
                onDragOver={(event) => event.preventDefault()}
                onDragLeave={(event) => { if (!event.currentTarget.contains(event.relatedTarget as Node)) setConsolidationDragging(false); }}
                onDrop={handleConsolidationDrop}
                onClick={() => !consolidationLoading && consolidationInputRef.current?.click()}
                role="button"
                tabIndex={0}
                onKeyDown={(event) => {
                  if (event.key === "Enter" || event.key === " ") {
                    event.preventDefault();
                    consolidationInputRef.current?.click();
                  }
                }}
              >
                <div className="upload-icon">
                  {consolidationLoading ? <LoaderCircle className="spin" size={23} /> : <Layers3 size={23} />}
                </div>
                <div className="upload-copy">
                  <strong>{consolidationLoading ? "Estamos uniendo tus hojas…" : consolidationWorkbook ? "Cargar otro Excel" : "Arrastrá tu Excel acá"}</strong>
                  <span>{consolidationLoading ? consolidationProgress : "o hacé clic para buscarlo en tu computadora"}</span>
                </div>
                {!consolidationLoading && (
                  <button className="browse-button" type="button">
                    Elegir archivo <ArrowRight size={15} />
                  </button>
                )}
                <input
                  ref={consolidationInputRef}
                  className="visually-hidden"
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={handleConsolidationFileChange}
                  aria-label="Subir Excel para unir sus hojas"
                />
              </div>
              <div className="upload-footnote">
                <LockKeyhole size={13} /> Hasta 50 MB y 1.048.576 filas en el archivo final; el tiempo depende de tu computadora.
              </div>
              {consolidationError && (
                <div className="error-message" role="alert">
                  <CircleHelp size={16} />{consolidationError}
                </div>
              )}
              {consolidationWorkbook && (
                <section className="consolidation-result" aria-live="polite">
                  <div className="consolidation-result-heading">
                    <span className="summary-file-icon"><FileSpreadsheet size={19} /></span>
                    <div className="summary-file-copy">
                      <strong>{consolidationWorkbook.name}</strong>
                      <span>
                        {consolidationWorkbook.rows.length.toLocaleString("es-AR")} filas reunidas de{" "}
                        {consolidationWorkbook.sheets.length.toLocaleString("es-AR")} hojas ·{" "}
                        {consolidationWorkbook.headers.length.toLocaleString("es-AR")} columnas
                      </span>
                    </div>
                    <span className="file-ready"><Check size={14} /> Listo</span>
                  </div>
                  <div className="consolidation-note">
                    <Layers3 size={15} />
                    <span>
                      Las columnas se alinean por encabezado; si una hoja no tiene una
                      columna, sus celdas quedan vacías. Se agrega “
                      {consolidationWorkbook.headers[0]}” para identificar de dónde salió
                      cada fila.
                    </span>
                  </div>
                  <button
                    className="export-button consolidation-download"
                    onClick={() => void exportConsolidatedWorkbook()}
                    type="button"
                    disabled={consolidationExporting}
                  >
                    {consolidationExporting ? <LoaderCircle className="spin" size={15} /> : <Download size={15} />}
                    {consolidationExporting ? consolidationProgress || "Preparando Excel…" : "Descargar Excel unificado"}
                  </button>
                </section>
              )}
            </section>
          ) : activePage === "top-products" ? (
            <section className="top-products-screen">
              <div className="welcome-topline"><span className="eyebrow-dot" /> NUEVA FUNCIONALIDAD</div>
              <h1>Los 5 productos</h1>
              <p className="welcome-description">
                Revisá por vendedor qué clientes cumplen el objetivo de los cinco
                productos y cuáles todavía no los están vendiendo.
              </p>
              <div
                className={`upload-card top-products-upload ${topProductsDragging ? "dragging" : ""} ${topProductsLoading ? "is-loading" : ""}`}
                onDragEnter={(event) => { event.preventDefault(); setTopProductsDragging(true); }}
                onDragOver={(event) => event.preventDefault()}
                onDragLeave={(event) => { if (!event.currentTarget.contains(event.relatedTarget as Node)) setTopProductsDragging(false); }}
                onDrop={handleTopProductsDrop}
                onClick={() => !topProductsLoading && topProductsInputRef.current?.click()}
                role="button"
                tabIndex={0}
                onKeyDown={(event) => {
                  if (event.key === "Enter" || event.key === " ") {
                    event.preventDefault();
                    topProductsInputRef.current?.click();
                  }
                }}
              >
                <div className="upload-icon">
                  {topProductsLoading ? <LoaderCircle className="spin" size={23} /> : <Upload size={23} />}
                </div>
                <div className="upload-copy">
                  <strong>{topProductsLoading ? "Estamos leyendo tu archivo…" : topProductsWorkbook ? "Cargar otro Excel" : "Arrastrá el Excel acá"}</strong>
                  <span>{topProductsLoading ? "Los datos se procesan en este dispositivo." : "o hacé clic para buscarlo en tu computadora"}</span>
                </div>
                {!topProductsLoading && (
                  <button className="browse-button" type="button">
                    Elegir archivo <ArrowRight size={15} />
                  </button>
                )}
                <input
                  ref={topProductsInputRef}
                  className="visually-hidden"
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={handleTopProductsFileChange}
                  aria-label="Subir Excel para Los 5 productos"
                />
              </div>
              <div className="upload-footnote">
                <LockKeyhole size={13} /> Excel con la base de clientes y las ventas, hasta 50 MB. Se procesa en este dispositivo.
              </div>
              {topProductsError && (
                <div className="error-message" role="alert">
                  <CircleHelp size={16} />{topProductsError}
                </div>
              )}
              {topProductsWorkbook && (
                <div className="top-products-file" role="status">
                  <span className="summary-file-icon"><FileSpreadsheet size={19} /></span>
                  <div className="summary-file-copy">
                    <strong>{topProductsWorkbook.name}</strong>
                    <span>
                      {topProductsWorkbook.rows.length.toLocaleString("es-AR")} clientes ·{" "}
                      {topProductsWorkbook.transactionCount.toLocaleString("es-AR")} ventas de los productos analizadas · hojas de clientes y ventas
                    </span>
                  </div>
                  <span className="file-ready"><Check size={14} /> Informe listo</span>
                </div>
              )}
              {topProductsWorkbook && (
                <section className="product-report">
                  <div className="product-report-heading">
                    <div>
                      <h2>Seguimiento por vendedor</h2>
                      <p>Un cliente cumple cuando alcanza las cantidades objetivo en cada producto.</p>
                    </div>
                    <button
                      className="export-button"
                      onClick={downloadTopProductsReport}
                      type="button"
                    >
                      <Download size={15} /> Descargar informe
                    </button>
                  </div>
                  <div className="product-threshold-note">
                    Objetivos: Café ≥ 6 unidades; Edulco, Azúcar, Yerba y Té ≥ 1 unidad.
                    {" "}{topProductsWorkbook.formulaResultRows.toLocaleString("es-AR")} clientes tienen los cinco resultados guardados en fórmulas; las celdas sin resultado se calculan desde la hoja de ventas.
                  </div>
                  <div className="product-report-stats">
                    <div><span>Clientes del vendedor</span><strong>{productReportStats.total.toLocaleString("es-AR")}</strong></div>
                    <div><span>Cumplen los 5</span><strong>{productReportStats.complete.toLocaleString("es-AR")}</strong></div>
                    <div><span>Vendieron, faltan productos</span><strong>{productReportStats.incomplete.toLocaleString("es-AR")}</strong></div>
                    <div><span>Sin ventas de estos 5</span><strong>{productReportStats.noSales.toLocaleString("es-AR")}</strong></div>
                  </div>
                  <div className="product-report-filters">
                    <label>
                      <span>Vendedor</span>
                      <OptionSelect
                        value={productSeller}
                        onChange={setProductSeller}
                        options={[
                          { value: "Todos", label: "Todos los vendedores" },
                          ...topProductsWorkbook.sellers.map((seller) => ({
                            value: seller,
                            label: seller,
                          })),
                        ]}
                        ariaLabel="Filtrar por vendedor"
                      />
                    </label>
                    <label>
                      <span>Estado</span>
                      <OptionSelect
                        value={productStatus}
                        onChange={(value) => {
                          if (isProductReportFilter(value)) {
                            setProductStatus(value);
                          }
                        }}
                        options={[
                          { value: "all", label: "Todos los clientes" },
                          { value: "complete", label: "Cumplen los 5" },
                          { value: "incomplete", label: "Les faltan productos" },
                          { value: "no-sales", label: "Sin ventas de los 5" },
                        ]}
                        ariaLabel="Filtrar por estado de cumplimiento"
                      />
                    </label>
                    <div className="product-search">
                      <Search size={15} aria-hidden="true" />
                      <input
                        value={productSearch}
                        onChange={(event) => setProductSearch(event.target.value)}
                        placeholder="Buscar cliente o código"
                        aria-label="Buscar cliente o código"
                      />
                    </div>
                    <span className="product-result-count">
                      {visibleProductRows.length.toLocaleString("es-AR")} clientes
                    </span>
                  </div>
                  <div className="table-wrapper product-report-table">
                    <table className="pivot-table">
                      <thead>
                        <tr>
                          <th>Código</th>
                          <th>Cliente</th>
                          <th>Vendedor</th>
                          {PRODUCT_REQUIREMENTS.map((product) => (
                            <th key={product.label}>{product.label}</th>
                          ))}
                          <th>Le faltan</th>
                          <th>Cumple 5/5</th>
                        </tr>
                      </thead>
                      <tbody>
                        {visibleProductRows.map((row, rowIndex) => (
                          <tr key={`${row.code}-${row.customer}-${row.seller}-${rowIndex}`}>
                            <td>{row.code || "—"}</td>
                            <td className="product-customer-cell">{row.customer}</td>
                            <td>{row.seller}</td>
                            {row.products.map((product, productIndex) => (
                              <td key={PRODUCT_REQUIREMENTS[productIndex].label}>
                                <span className={`product-result ${product.meetsRequirement ? "met" : "missed"}`}>
                                  {product.meetsRequirement ? (
                                    <Check size={15} aria-label="Cumple" />
                                  ) : (
                                    <X size={15} aria-label="No cumple" />
                                  )}
                                </span>
                              </td>
                            ))}
                            <td>{row.missingProducts.join(", ") || "—"}</td>
                            <td>
                              <span className={`product-compliance ${row.meetsAll ? "met" : "missed"}`}>
                                {row.meetsAll ? "Sí" : "No"}
                              </span>
                            </td>
                          </tr>
                        ))}
                        {visibleProductRows.length === 0 && (
                          <tr>
                            <td className="product-empty" colSpan={PRODUCT_REQUIREMENTS.length + 5}>
                              No hay clientes para los filtros seleccionados.
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </section>
              )}
            </section>
          ) : !workbook ? (
            <section className="welcome-screen">
              <div className="welcome-topline"><span className="eyebrow-dot" /> UNA FORMA MÁS CLARA DE VER TUS VENTAS</div>
              <h1>Los números cuentan<br /><span>historias.</span></h1>
              <p className="welcome-description">Subí tu Excel y descubrí qué está pasando con tus ventas.<br />Sin fórmulas, sin vueltas y con tus datos siempre privados.</p>

              <div
                className={`upload-card ${dragging ? "dragging" : ""} ${loading ? "is-loading" : ""}`}
                onDragEnter={(event) => { event.preventDefault(); setDragging(true); }}
                onDragOver={(event) => event.preventDefault()}
                onDragLeave={(event) => { if (!event.currentTarget.contains(event.relatedTarget as Node)) setDragging(false); }}
                onDrop={handleDrop}
                onClick={() => !loading && inputRef.current?.click()}
                role="button"
                tabIndex={0}
                onKeyDown={(event) => {
                  if (event.key === "Enter" || event.key === " ") {
                    event.preventDefault();
                    inputRef.current?.click();
                  }
                }}
              >
                <div className="upload-icon">{loading ? <LoaderCircle className="spin" size={23} /> : <Upload size={23} />}</div>
                <div className="upload-copy">
                  <strong>{loading ? "Estamos leyendo tu archivo…" : "Arrastrá tu Excel acá"}</strong>
                  <span>{loading ? "Los datos se procesan en este dispositivo." : "o hacé clic para buscarlo en tu computadora"}</span>
                </div>
                {!loading && <button className="browse-button" type="button">Elegir archivo <ArrowRight size={15} /></button>}
                <input
                  ref={inputRef}
                  className="visually-hidden"
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  onChange={handleFileChange}
                  aria-label="Subir un archivo Excel"
                />
              </div>
              <div className="upload-footnote"><LockKeyhole size={13} /> Excel o CSV, hasta 50 MB. No guardamos tu archivo.</div>
              {error && <div className="error-message" role="alert"><CircleHelp size={16} />{error}</div>}
              <div className="welcome-divider"><span>QUÉ VAS A PODER HACER</span></div>
              <div className="feature-grid">
                <div className="feature-card">
                  <span className="feature-icon green"><Rows3 size={18} /></span>
                  <strong>Armá tu tabla</strong>
                  <p>Cruzá meses, productos, vendedores y lo que necesites.</p>
                </div>
                <div className="feature-card">
                  <span className="feature-icon gold"><ArrowDownUp size={18} /></span>
                  <strong>Seguí el evolutivo</strong>
                  <p>Encontrá tendencias en ventas, kilos o unidades.</p>
                </div>
                <div className="feature-card">
                  <span className="feature-icon blue"><Sparkles size={18} /></span>
                  <strong>Sin limpiar a mano</strong>
                  <p>Organizamos el archivo y te avisamos si algo no cierra.</p>
                </div>
              </div>
              <div className="welcome-bottom"><span>HECHO PARA EMPRESAS QUE CONOCEN CADA GRANO.</span><span>COMPARA © 2026</span></div>
            </section>
          ) : (
            <div className="dashboard">
              <div className="dashboard-heading">
                <div>
                  <div className="welcome-topline"><span className="eyebrow-dot" /> TU NEGOCIO, EN PERSPECTIVA</div>
                  <h1>{activePage === "overview" ? "La película completa." : "Miremos los datos."}</h1>
                  <p>Explorá el detalle y encontrá nuevas oportunidades de crecimiento.</p>
                </div>
                <button className="outline-button" onClick={() => inputRef.current?.click()} type="button"><RefreshCw size={15} /> Cambiar archivo</button>
              </div>

              <div className="file-summary">
                <span className="summary-file-icon"><FileSpreadsheet size={19} /></span>
                <div className="summary-file-copy"><strong>{workbook.name}</strong><span>{analysisWorkbook?.rows.length.toLocaleString("es-AR")} registros analizados · {workbook.headers.length} columnas{comparisonSources.length ? ` · ${comparisonSources.length + 1} archivos` : ""}</span></div>
                <div className="summary-divider" />
                <span className="sheet-label">Hoja</span>
                <OptionSelect
                  value={workbook.sheet}
                  onChange={handleSheetChange}
                  options={workbook.sheets.map((sheet) => ({ value: sheet, label: sheet }))}
                  ariaLabel="Elegir hoja del libro"
                />
                <span className="file-ready"><Check size={14} /> Listo para analizar</span>
              </div>

              <div className="compare-files-panel">
                <div className="compare-files-heading">
                  <div>
                    <strong>Comparar archivos</strong>
                    <span>
                      {comparisonSources.length
                        ? `${comparisonSources.length + 1} archivos cargados`
                        : "Agregá uno o más Excel para compararlos con el archivo base"}
                    </span>
                  </div>
                  <button
                    className="outline-button"
                    onClick={() => comparisonInputRef.current?.click()}
                    type="button"
                  >
                    <Plus size={15} /> Agregar archivos
                  </button>
                  <input
                    ref={comparisonInputRef}
                    className="visually-hidden"
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    multiple
                    onChange={handleComparisonFiles}
                    aria-label="Agregar archivos para comparar"
                  />
                </div>
                {comparisonError && (
                  <div className="error-message" role="alert">
                    <CircleHelp size={16} />{comparisonError}
                  </div>
                )}
                {comparisonSources.length > 0 && (
                  <>
                    <div className="compare-mode-control" role="group" aria-label="Modo de comparación">
                      <span>Mostrar resultados</span>
                      <button
                        className={comparisonMode === "periods" ? "selected" : ""}
                        onClick={() => setComparisonMode("periods")}
                        type="button"
                      >
                        En columnas por archivo
                      </button>
                      <button
                        className={comparisonMode === "separate" ? "selected" : ""}
                        onClick={() => setComparisonMode("separate")}
                        type="button"
                      >
                        Un archivo por vez
                      </button>
                    </div>
                    <div className="compare-source-list">
                      <div className="compare-source base-source">
                        <FileSpreadsheet size={16} />
                        <span>{workbook.name}</span>
                        <b>Base</b>
                      </div>
                      {comparisonSources.map((source) => (
                        <details className="compare-source-mapping" key={source.id}>
                          <summary className="compare-source">
                            <FileSpreadsheet size={16} />
                            <input
                              defaultValue={source.label}
                              maxLength={60}
                              aria-label={`Nombre del período ${source.name}`}
                              onClick={(event) => event.stopPropagation()}
                              onBlur={(event) =>
                                updateSourceLabel(source.id, event.currentTarget.value.trim())
                              }
                            />
                            <span>
                              {Object.values(source.fieldMapping).filter(Boolean).length}/
                              {workbook.headers.length} columnas mapeadas
                            </span>
                            <button
                              type="button"
                              className="remove-source-button"
                              aria-label={`Quitar ${source.name}`}
                              title="Quitar archivo"
                              onClick={(event) => {
                                event.preventDefault();
                                event.stopPropagation();
                                removeComparisonSource(source.id);
                              }}
                            >
                              <X size={15} />
                            </button>
                          </summary>
                          <div className="field-mapping-grid">
                            {workbook.headers.map((header) => (
                              <label className="field-mapping-row" key={header}>
                                <span>{header}</span>
                                <OptionSelect
                                  value={source.fieldMapping[header] ?? ""}
                                  onChange={(mappedHeader) =>
                                    updateSourceMapping(
                                      source.id,
                                      header,
                                      mappedHeader,
                                    )
                                  }
                                  options={selectorOptions(source.headers)}
                                  placeholder="No mapear"
                                  ariaLabel={`Columna de ${source.name} que corresponde a ${header}`}
                                />
                              </label>
                            ))}
                          </div>
                        </details>
                      ))}
                    </div>
                    {comparisonMode === "separate" && (
                      <div className="compare-source-tabs" role="tablist" aria-label="Archivo para ver">
                        <button
                          className={activeComparisonSource === "base" ? "selected" : ""}
                          onClick={() => setActiveComparisonSource("base")}
                          type="button"
                        >
                          {workbook.name.replace(/\.[^.]+$/, "")}
                        </button>
                        {comparisonSources.map((source) => (
                          <button
                            className={activeComparisonSource === source.id ? "selected" : ""}
                            key={source.id}
                            onClick={() => setActiveComparisonSource(source.id)}
                            type="button"
                          >
                            {source.label}
                          </button>
                        ))}
                      </div>
                    )}
                  </>
                )}
              </div>

              {error && <div className="error-message" role="alert"><CircleHelp size={16} />{error}</div>}
              {successMessage && <div className="success-message" role="status"><Check size={16} />{successMessage}</div>}
              {dateWarnings > 0 && (
                <div className="data-warning" role="status">
                  <CircleHelp size={15} />
                  <span>
                    Detectamos una diferencia entre <b>Fecha</b> y <b>Mes</b> en{" "}
                    <b>{dateWarnings.toLocaleString("es-AR")} filas</b>. El evolutivo usa
                    la columna Fecha.
                  </span>
                </div>
              )}

              <div className="metrics-grid">
                <div className="metric-card">
                  <span className="metric-label">REGISTROS EN EL ARCHIVO</span>
                  <span className="metric-value">{analysisWorkbook?.rows.length.toLocaleString("es-AR")}</span>
                  <span className="metric-note"><Layers3 size={14} /> {comparisonSources.length ? "Filas de todos los archivos" : "Todas las filas incluidas"}</span>
                  <span className="metric-decoration rows-decoration"><Rows3 size={45} /></span>
                </div>
                <div className="metric-card">
                  <span className="metric-label">VENDEDORES / CANALES</span>
                  <span className="metric-value">{formatNumber(quickMetrics.sellers)}</span>
                  <span className="metric-note"><UsersRound size={14} /> Valores únicos detectados</span>
                  <span className="metric-decoration seller-decoration"><UsersRound size={45} /></span>
                </div>
                <div className="metric-card">
                  <span className="metric-label">PRODUCTOS</span>
                  <span className="metric-value">{formatNumber(quickMetrics.products)}</span>
                  <span className="metric-note"><Coffee size={14} /> En tu catálogo</span>
                  <span className="metric-decoration product-decoration"><Coffee size={45} /></span>
                </div>
                <div className="metric-card">
                  <span className="metric-label">CLIENTES</span>
                  <span className="metric-value">{formatNumber(quickMetrics.customers)}</span>
                  <span className="metric-note"><UsersRound size={14} /> En este período</span>
                  <span className="metric-decoration customer-decoration"><ArrowUpRight size={45} /></span>
                </div>
              </div>

              <section className="pivot-section">
                <div className="section-heading">
                  <div><span className="section-kicker">EXPLORADOR</span><h2>Compará tus ventas</h2><p>Cruzá meses, vendedores y productos para ver las diferencias.</p></div>
                  {pivot && <button className="export-button" onClick={exportPivot} type="button"><ArrowDownToLine size={15} /> Exportar Excel</button>}
                </div>

                <div className="comparison-presets" aria-label="Análisis rápidos">
                  <span>EMPEZAR CON</span>
                  {sellerField && (
                    <button
                      className={rowField === sellerField && columnFields.length === 1 && columnFields[0] === (monthField || dateField) ? "selected" : ""}
                      onClick={() => applyComparisonPreset("sellers")}
                      type="button"
                    >
                      <UsersRound size={14} /> Vendedores por mes
                    </button>
                  )}
                  {productField && (
                    <button
                      className={rowField === productField && columnFields.length === 1 && columnFields[0] === (monthField || dateField) ? "selected" : ""}
                      onClick={() => applyComparisonPreset("products")}
                      type="button"
                    >
                      <Coffee size={14} /> Productos por mes
                    </button>
                  )}
                  <button
                    className={rowField === dateField && columnFields.length === 0 ? "selected" : ""}
                    onClick={() => applyComparisonPreset("evolution")}
                    type="button"
                  >
                    <BarChart3 size={14} /> Evolutivo mensual
                  </button>
                  {savedComparisons.map((comparison) => (
                    <span className="saved-comparison" key={comparison.id}>
                      <button
                        className="saved-comparison-apply"
                        onClick={() => applySavedComparison(comparison)}
                        type="button"
                        title={`Aplicar ${comparison.name}`}
                      >
                        <BookmarkPlus size={14} /> {comparison.name}
                      </button>
                      <button
                        className="saved-comparison-remove"
                        onClick={() => removeSavedComparison(comparison.id)}
                        type="button"
                        aria-label={`Eliminar comparativa ${comparison.name}`}
                        title="Eliminar comparativa guardada"
                      >
                        <X size={13} />
                      </button>
                    </span>
                  ))}
                  <button
                    className="save-comparison-button"
                    onClick={openSaveComparisonDialog}
                    type="button"
                    disabled={!activeMetricFields.length}
                  >
                    <Plus size={14} /> Guardar comparativa
                  </button>
                </div>

                <div className="pivot-controls">
                  <div className="pivot-control">
                    <span className="control-label"><Rows3 size={14} /> FILAS</span>
                    <OptionSelect value={rowField} onChange={updateRowField} options={selectorOptions(workbook.headers.filter((field) => !columnFields.includes(field)))} ariaLabel="Campo de filas" />
                  </div>
                  <div className="pivot-control">
                    <span className="control-label"><ColumnsIcon /> COLUMNAS</span>
                    <MultiSelectFilter
                      options={workbook.headers.filter((field) => field !== rowField)}
                      selectedValues={columnFields}
                      onChange={updateColumnFields}
                      label="Seleccionar campos de columnas"
                      allowSelectAll={false}
                      selectionSummary={(fields) =>
                        fields.length === 0
                          ? "Sin columnas"
                          : fields.length === 1
                            ? fields[0]
                            : `${fields.length} campos seleccionados`
                      }
                    />
                    <span className="metric-selection-note">
                      Los campos numéricos seleccionados aparecen en columnas separadas.
                    </span>
                  </div>
                  {columnMetricFields.length > 0 && (
                    <div className="pivot-control">
                      <span className="control-label"><BarChart3 size={14} /> VALOR</span>
                      <span className="metric-selection-note">
                        Se resumen: {columnMetricFields.join(", ")}
                      </span>
                    </div>
                  )}
                  {columnMetricFields.length === 0 && (
                    <div className="pivot-control">
                      <span className="control-label"><BarChart3 size={14} /> VALOR</span>
                      <MultiSelectFilter
                        options={numericFields}
                        selectedValues={metricFields}
                        onChange={updateMetricFields}
                        label="Seleccionar valores numéricos"
                        selectionSummary={(fields) =>
                          fields.length === 0
                            ? "Elegir valor"
                            : fields.length === 1
                              ? fields[0]
                              : `${fields.length} valores`
                        }
                      />
                      <span className="metric-selection-note">
                        Elegí qué valores numéricos resumir.
                      </span>
                    </div>
                  )}
                  <div className="pivot-control">
                    <span className="control-label"><ArrowDownUp size={14} /> RESUMIR POR</span>
                    <OptionSelect value={metricMode} onChange={(value) => setMetricMode(value as MetricMode)} options={[
                      { value: "sum", label: "Suma" },
                      { value: "average", label: "Promedio" },
                      { value: "count", label: "Recuento" },
                    ]} ariaLabel="Cómo resumir valores" />
                  </div>
                </div>

                {(rowField === dateField || columnFields.includes(dateField)) && (
                  <div className="date-grouping">
                    <span>Agrupar fecha en filas o columnas por</span>
                    <div className="segmented-control">
                      {([
                        ["month", "Mes"],
                        ["quarter", "Trimestre"],
                        ["year", "Año"],
                        ["day", "Día"],
                      ] as const).map(([value, label]) => (
                        <button
                          className={dateGrouping === value ? "selected" : ""}
                          key={value}
                          onClick={() => setDateGrouping(value)}
                          type="button"
                        >{label}</button>
                      ))}
                    </div>
                  </div>
                )}

                <div className="filters-row">
                  <div className="filter-title"><Filter size={15} /><span>FILTROS</span></div>
                  <div className="filters-list">
                    {filters.map((filter, index) => (
                      <div className="dimension-filter" key={`${index}-${filter.field}`}>
                        <OptionSelect
                          value={filter.field}
                          onChange={(field) =>
                            updateDimensionFilter(index, { field, values: [] })
                          }
                          options={selectorOptions(
                            workbook.headers.filter(
                              (field) =>
                                !filters.some(
                                  (other, otherIndex) =>
                                    otherIndex !== index && other.field === field,
                                ),
                            ),
                          )}
                          placeholder="Elegir columna"
                          ariaLabel={`Columna de filtro ${index + 1}`}
                        />
                        {filter.field && (
                          <MultiSelectFilter
                            options={filterOptionsByField[filter.field] ?? []}
                            selectedValues={filter.values}
                            onChange={(values) =>
                              updateDimensionFilter(index, { ...filter, values })
                            }
                            label={`Filtrar por ${filter.field}`}
                          />
                        )}
                        <button
                          className="remove-dimension-filter"
                          onClick={() => removeDimensionFilter(index)}
                          type="button"
                          aria-label={`Quitar filtro ${filter.field || index + 1}`}
                          title="Quitar filtro"
                        >
                          <X size={14} />
                        </button>
                      </div>
                    ))}
                  </div>
                  {workbook.headers.some(
                    (field) => !filters.some((filter) => filter.field === field),
                  ) && (
                    <button
                      className="add-dimension-filter"
                      onClick={() => addDimensionFilter()}
                      type="button"
                    >
                      <Plus size={14} /> Agregar filtro
                    </button>
                  )}
                  {sellerField && (
                    !filters.some((filter) => filter.field === sellerField) && (
                      <button className="quick-filter-button" onClick={() => addDimensionFilter(sellerField)} type="button">
                        <UsersRound size={14} /> Agregar vendedor
                      </button>
                    )
                  )}
                  <span className="filtered-count">{filteredRows.length.toLocaleString("es-AR")} filas</span>
                </div>

                <p className="comparison-hint">
                  Para comparar agosto con septiembre, dejá <b>Mes</b> en columnas y
                  marcá ambos meses en el filtro. Usá las filas para ver cada vendedor o producto;
                  también aparece la variación y el cambio porcentual.
                </p>

                {!rowField || activeMetricFields.length === 0 ? (
                  <div className="empty-pivot">
                    <span className="empty-pivot-icon"><Settings2 size={20} /></span>
                    <strong>Elegí las columnas del reporte</strong>
                    <span>Seleccioná un campo para las filas y una o varias métricas numéricas.</span>
                  </div>
                ) : pivot && pivot.rowLabels.length > 0 ? (
                  <>
                    <div className="pivot-summary">
                      <div><span>RESULTADO DEL ANÁLISIS</span><strong>{activeMetricFields.map((field) => {
                        const total = pivot.grandTotals.get(field) ?? { total: 0, count: 0 };
                        return `${field}: ${formatPivotMetric(field, pivotCellValue(total, metricMode))}`;
                      }).join(" · ")}</strong></div>
                      <span className="summary-caption">{metricMode === "sum" ? "Suma" : metricMode === "average" ? "Promedio por fila" : "Registros"} de <b>{activeMetricFields.join(", ")}</b>{activeFilterSummary.length > 0 ? <> · Filtros: <b>{activeFilterSummary.join(" · ")}</b></> : null}</span>
                    </div>
                    {rowField === dateField && chartItems.length > 0 && (
                      <div className="chart-card">
                        <div className="chart-heading"><div><strong>Evolución de {metricField}</strong><span>Por {dateGrouping === "month" ? "mes" : dateGrouping === "quarter" ? "trimestre" : dateGrouping === "year" ? "año" : "día"}</span></div><span className="chart-period"><span /> {chartItems.length} períodos</span></div>
                        <div className="bar-chart" role="img" aria-label={`Gráfico de ${metricField} por período`}>
                          {chartItems.map((item, index) => (
                            <div className="bar-column" key={item.label} title={`${item.label}: ${formatPivotMetric(metricField, item.value)}`}>
                              <div className="bar-value">{item.value === 0 ? "" : formatPivotMetric(metricField, item.value)}</div>
                              <div className="bar-track"><div className="bar-fill" style={{ height: `${Math.max(3, (Math.abs(item.value) / maximumChartValue) * 100)}%`, backgroundColor: PALETTE[index % PALETTE.length] }} /></div>
                              <div className="bar-label">{item.label.split(" ")[0].slice(0, 3)}</div>
                            </div>
                          ))}
                        </div>
                      </div>
                    )}
                    <div className="table-wrapper">
                      <table className="pivot-table">
                        <thead>
                          <tr>
                            <th>{rowField}</th>
                            {valueColumns.map(({ key, period, field }) => (
                              <th className="numeric-cell" key={key}>
                                {pivotColumnDimensionFields.length
                                  ? `${pivot.columnDisplayLabels.get(period) ?? period} · ${field}`
                                  : field}
                              </th>
                            ))}
                            {comparisonPeriods.length === 2 &&
                              activeMetricFields.map((field) => (
                                <Fragment key={`comparison-${field}`}>
                                  <th className="numeric-cell comparison-column">{field} · Variación</th>
                                  <th className="numeric-cell comparison-column">{field} · Cambio %</th>
                                </Fragment>
                              ))}
                            {activeMetricFields.map((field) => (
                              <th className="numeric-cell total-column" key={`total-${field}`}>
                                {field} · Total
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {pivot.rowLabels.map((label) => {
                            const rowValues = pivot.rowTotals.get(label);
                            return (
                              <tr key={label}>
                                <td className="row-label">{label}</td>
                                {valueColumns.map(({ key, period, field }) => {
                                  const cell = rowValues?.get(period)?.get(field);
                                  return (
                                    <td className="numeric-cell" key={key}>
                                      {!cell ? <span className="dash-cell">—</span> : formatPivotMetric(field, pivotCellValue(cell, metricMode))}
                                    </td>
                                  );
                                })}
                                {comparisonPeriods.length === 2 &&
                                  activeMetricFields.map((field) => {
                                    const values = comparisonPeriods.map((period) =>
                                      pivotCellValue(rowValues?.get(period)?.get(field), metricMode),
                                    );
                                    const difference = values[1] - values[0];
                                    const percentChange =
                                      values[0] === 0 ? null : (difference / Math.abs(values[0])) * 100;
                                    const changeClass =
                                      difference > 0 ? "positive-change" : difference < 0 ? "negative-change" : "";
                                    return (
                                      <Fragment key={`comparison-${field}`}>
                                        <td className={`numeric-cell comparison-column ${changeClass}`}>
                                          {formatPivotMetric(field, difference)}
                                        </td>
                                        <td className={`numeric-cell comparison-column ${changeClass}`}>
                                          {percentChange === null ? "—" : `${formatNumber(percentChange, 1)}%`}
                                        </td>
                                      </Fragment>
                                    );
                                  })}
                                {activeMetricFields.map((field) => {
                                  const total = metricTotal(
                                    pivot.columnLabels.map((period) =>
                                      rowValues?.get(period)?.get(field),
                                    ),
                                  );
                                  return (
                                    <td className="numeric-cell total-column" key={`total-${field}`}>
                                      {formatPivotMetric(field, pivotCellValue(total, metricMode))}
                                    </td>
                                  );
                                })}
                              </tr>
                            );
                          })}
                        </tbody>
                        <tfoot>
                          <tr>
                            <th>Total general</th>
                            {valueColumns.map(({ key, period, field }) => (
                              <td className="numeric-cell" key={key}>
                                {formatPivotMetric(
                                  field,
                                  pivotCellValue(pivot.totals.get(period)?.get(field), metricMode),
                                )}
                              </td>
                            ))}
                            {comparisonPeriods.length === 2 &&
                              activeMetricFields.map((field) => {
                                const values = comparisonPeriods.map((period) =>
                                  pivotCellValue(pivot.totals.get(period)?.get(field), metricMode),
                                );
                                const difference = values[1] - values[0];
                                const percentChange =
                                  values[0] === 0 ? null : (difference / Math.abs(values[0])) * 100;
                                const changeClass =
                                  difference > 0 ? "positive-change" : difference < 0 ? "negative-change" : "";
                                return (
                                  <Fragment key={`comparison-total-${field}`}>
                                    <td className={`numeric-cell comparison-column ${changeClass}`}>
                                      {formatPivotMetric(field, difference)}
                                    </td>
                                    <td className={`numeric-cell comparison-column ${changeClass}`}>
                                      {percentChange === null ? "—" : `${formatNumber(percentChange, 1)}%`}
                                    </td>
                                  </Fragment>
                                );
                              })}
                            {activeMetricFields.map((field) => (
                              <td className="numeric-cell total-column" key={`total-${field}`}>
                                {formatPivotMetric(
                                  field,
                                  pivotCellValue(pivot.grandTotals.get(field), metricMode),
                                )}
                              </td>
                            ))}
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </>
                ) : (
                  <div className="empty-pivot">
                    <span className="empty-pivot-icon"><Search size={20} /></span>
                    <strong>No encontramos datos para esta combinación</strong>
                    <span>Probá elegir otro campo o quitar los filtros activos.</span>
                  </div>
                )}
              </section>

              {activePage === "data" && (
                <section className="data-section">
                  <div className="section-heading">
                    <div><span className="section-kicker">ARCHIVO ORIGINAL</span><h2>Explorar filas</h2><p>Buscá y revisá una muestra de los datos importados.</p></div>
                  </div>
                  <div className="data-search"><Search size={16} /><input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Buscar en todas las columnas…" aria-label="Buscar en los datos" /><span>{shownRows.length.toLocaleString("es-AR")} resultados</span></div>
                  <div className="table-wrapper data-table-wrapper">
                    <table className="pivot-table data-table">
                      <thead><tr>{analysisWorkbook?.headers.map((header) => <th key={header}>{header}</th>)}</tr></thead>
                      <tbody>{shownRows.slice(0, 50).map((row, index) => <tr key={index}>{analysisWorkbook?.headers.map((header) => {
                        const value = row[header] ?? null;
                        const displayValue =
                          header === comparisonSourceField && typeof value === "string"
                            ? value === "base"
                              ? workbook.name.replace(/\.[^.]+$/, "")
                              : comparisonSources.find((source) => source.id === value)?.label ?? value
                            : cellText(value);
                        return <td key={header}>{displayValue}</td>;
                      })}</tr>)}</tbody>
                    </table>
                  </div>
                  {shownRows.length > 50 && <div className="table-note">Mostrando las primeras 50 filas de {shownRows.length.toLocaleString("es-AR")} resultados.</div>}
                </section>
              )}

              <div className="dashboard-footer"><span><LockKeyhole size={13} /> Tu archivo se procesa solo en este navegador.</span><button onClick={() => { setWorkbook(null); setError(""); }} type="button">Empezar de nuevo <ArrowRight size={14} /></button></div>
              <input ref={inputRef} className="visually-hidden" type="file" accept=".xlsx,.xls,.csv" onChange={handleFileChange} aria-label="Cambiar archivo Excel" />
            </div>
          )}
        </div>
      </section>
      {saveDialogOpen && (
        <div
          className="save-dialog-backdrop"
          onKeyDown={(event) => {
            if (event.key === "Escape") setSaveDialogOpen(false);
          }}
          onClick={(event) => {
            if (event.target === event.currentTarget) setSaveDialogOpen(false);
          }}
        >
          <form
            className="save-dialog"
            role="dialog"
            aria-modal="true"
            aria-labelledby="save-dialog-title"
            onSubmit={saveCurrentComparison}
          >
            <span className="save-dialog-icon"><BookmarkPlus size={19} /></span>
            <h2 id="save-dialog-title">Guardar comparativa</h2>
            <p>Escribí un nombre para encontrarla después en “Empezar con”.</p>
            <label htmlFor="comparison-name">Nombre</label>
            <input
              id="comparison-name"
              autoFocus
              maxLength={60}
              value={saveComparisonName}
              onChange={(event) => setSaveComparisonName(event.target.value)}
              placeholder="Ej.: Ventas por cliente y producto"
              required
            />
            <div className="save-dialog-actions">
              <button
                type="button"
                className="save-dialog-cancel"
                onClick={() => setSaveDialogOpen(false)}
              >
                Cancelar
              </button>
              <button
                type="submit"
                className="save-dialog-confirm"
                disabled={!saveComparisonName.trim()}
              >
                Guardar
              </button>
            </div>
          </form>
        </div>
      )}
    </main>
  );
}

function ColumnsIcon() {
  return <span className="columns-icon"><span /><span /><span /></span>;
}
