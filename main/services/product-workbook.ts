import ExcelJS from 'exceljs';
import { randomUUID } from 'node:crypto';
import { getDatabase, generateShortId, now } from '../db';
import { adjustProductStock } from './inventory';
import { canonicalizeLocalizedAmount, type RegionalSnapshot } from '../countries';

type Db = ReturnType<typeof getDatabase>;

export const PRODUCTS_SHEET = 'Productos';
export const CATEGORIES_SHEET = 'Categorias';
const TEMPLATE_ROWS = 500;
const MAX_IMPORT_ROWS = 5000;

const CATEGORY_COLORS: Record<string, string> = {
  rojo: 'red',
  naranja: 'orange',
  ambar: 'amber',
  amarillo: 'yellow',
  lima: 'lime',
  verde: 'green',
  esmeralda: 'emerald',
  'verde azulado': 'teal',
  cian: 'cyan',
  celeste: 'sky',
  azul: 'blue',
  indigo: 'indigo',
  violeta: 'violet',
  morado: 'purple',
  fucsia: 'fuchsia',
  rosado: 'pink',
  rosa: 'rose',
};
const COLOR_LABEL_BY_KEY = Object.fromEntries(
  Object.entries(CATEGORY_COLORS).map(([label, key]) => [key, label]),
);

const SALE_UNITS: Record<string, string> = {
  unidad: 'each',
  und: 'each',
  each: 'each',
  kg: 'kg',
  g: 'g',
  lb: 'lb',
  l: 'l',
  ml: 'ml',
  cl: 'cl',
  oz: 'oz',
};
const UNIT_LABELS = ['Unidad', 'kg', 'g', 'lb', 'l', 'ml'];

type ProductField =
  | 'name'
  | 'category'
  | 'cost'
  | 'margin'
  | 'price'
  | 'barcode'
  | 'sku'
  | 'track_inventory'
  | 'stock'
  | 'low_stock'
  | 'unit'
  | 'description'
  | 'active';

const PRODUCT_COLUMNS: Array<{
  field: ProductField;
  header: string;
  width: number;
  note: string;
  text?: boolean;
}> = [
  { field: 'name', header: 'Nombre *', width: 32, note: 'Nombre que se ve en el POS.' },
  {
    field: 'category',
    header: 'Categoría *',
    width: 22,
    note: 'Elige de la lista (hoja Categorias). Si escribes una nueva, se crea sola.',
  },
  { field: 'cost', header: 'Costo', width: 12, note: 'Lo que te cuesta el producto.' },
  {
    field: 'margin',
    header: 'Margen %',
    width: 11,
    note: 'Ganancia sobre el costo. Con Costo y Margen, el precio se calcula solo (redondeado a 100).',
  },
  {
    field: 'price',
    header: 'Precio venta *',
    width: 15,
    note: 'Se calcula con Costo y Margen, o escríbelo directamente.',
  },
  {
    field: 'barcode',
    header: 'Código de barras',
    width: 20,
    note: 'Escanéalo aquí con el lector. Debe ser único.',
    text: true,
  },
  { field: 'sku', header: 'Referencia (SKU)', width: 16, note: 'Opcional.', text: true },
  {
    field: 'track_inventory',
    header: 'Controlar inventario',
    width: 12,
    note: 'Sí o No. Vacío = Sí cuando hay stock.',
  },
  { field: 'stock', header: 'Stock', width: 9, note: 'Cantidad disponible hoy.' },
  {
    field: 'low_stock',
    header: 'Stock mínimo',
    width: 10,
    note: 'Avisa cuando quede esta cantidad o menos.',
  },
  { field: 'unit', header: 'Unidad', width: 10, note: 'Vacío = Unidad.' },
  { field: 'description', header: 'Descripción', width: 36, note: 'Opcional.' },
  { field: 'active', header: 'Activo', width: 8, note: 'Vacío = Sí.' },
];

function normalizeHeader(value: string): string {
  return value
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/[*()]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

const PRODUCT_HEADER_ALIASES: Record<string, ProductField> = {
  nombre: 'name',
  name: 'name',
  categoria: 'category',
  category: 'category',
  costo: 'cost',
  cost: 'cost',
  'margen %': 'margin',
  margen: 'margin',
  margin: 'margin',
  'precio venta': 'price',
  precio: 'price',
  price: 'price',
  'codigo de barras': 'barcode',
  barcode: 'barcode',
  'referencia sku': 'sku',
  referencia: 'sku',
  sku: 'sku',
  'controlar inventario': 'track_inventory',
  track_inventory: 'track_inventory',
  stock: 'stock',
  'stock inicial': 'stock',
  'stock minimo': 'low_stock',
  low_stock_threshold: 'low_stock',
  unidad: 'unit',
  description: 'description',
  descripcion: 'description',
  activo: 'active',
  is_active: 'active',
};

function cellText(cell: ExcelJS.Cell): string {
  const value = cell.value as unknown;
  if (value === null || value === undefined) return '';
  if (typeof value === 'number') return Number.isInteger(value) ? value.toFixed(0) : String(value);
  if (typeof value === 'string') return value.trim();
  if (typeof value === 'boolean') return value ? 'sí' : 'no';
  if (value instanceof Date) return value.toISOString();
  if (typeof value === 'object') {
    const v = value as {
      richText?: Array<{ text: string }>;
      result?: unknown;
      text?: unknown;
      formula?: string;
      sharedFormula?: string;
    };
    if (Array.isArray(v.richText))
      return v.richText
        .map((part) => part.text)
        .join('')
        .trim();
    if ('formula' in v || 'sharedFormula' in v) {
      const result = v.result;
      if (result === null || result === undefined || typeof result === 'object') return '';
      if (typeof result === 'number')
        return Number.isInteger(result) ? result.toFixed(0) : String(result);
      return String(result).trim();
    }
    if (v.text !== undefined) return String(v.text).trim();
  }
  return String(value).trim();
}

function cellNumber(
  cell: ExcelJS.Cell,
  regional: RegionalSnapshot,
): { empty: true } | { empty: false; value: number | null } {
  const raw = cell.value as unknown;
  const formulaResult =
    raw && typeof raw === 'object' && ('formula' in raw || 'sharedFormula' in raw)
      ? (raw as { result?: unknown }).result
      : raw;
  if (typeof formulaResult === 'number')
    return { empty: false, value: Number.isFinite(formulaResult) ? formulaResult : null };
  const text = cellText(cell).replace(/\s+/g, '');
  if (!text) return { empty: true };
  const canonical = canonicalizeLocalizedAmount(text, regional);
  const parsed = canonical === null || canonical === '' ? NaN : Number(canonical);
  return { empty: false, value: Number.isFinite(parsed) ? parsed : null };
}

function parseYesNo(text: string): boolean | null {
  const norm = normalizeHeader(text);
  if (!norm) return null;
  if (['si', 's', 'yes', 'y', 'true', '1', 'x'].includes(norm)) return true;
  if (['no', 'n', 'false', '0'].includes(norm)) return false;
  return null;
}

function slugify(name: string): string {
  return name
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/(^-|-$)/g, '');
}

// ─── Template ────────────────────────────────────────────────────────────────

export async function buildProductTemplate(db: Db): Promise<Buffer> {
  const categories = db
    .prepare(
      `SELECT name, color, icon FROM categories
       WHERE deleted_at IS NULL AND is_active = 1
       ORDER BY CASE WHEN parent_id IS NULL THEN 0 ELSE 1 END, sort_order, name`,
    )
    .all() as Array<{ name: string; color: string | null; icon: string | null }>;

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Cantaro POS';

  const headerFill: ExcelJS.Fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFB5521D' },
  };
  const headerFont: Partial<ExcelJS.Font> = { bold: true, color: { argb: 'FFFFFFFF' } };

  const products = wb.addWorksheet(PRODUCTS_SHEET, {
    views: [{ state: 'frozen', ySplit: 1 }],
  });
  const catSheet = wb.addWorksheet(CATEGORIES_SHEET, {
    views: [{ state: 'frozen', ySplit: 1 }],
  });
  const lists = wb.addWorksheet('Listas', { state: 'veryHidden' });
  const help = wb.addWorksheet('Instrucciones');

  products.columns = PRODUCT_COLUMNS.map((col) => ({
    header: col.header,
    key: col.field,
    width: col.width,
    style: col.text ? { numFmt: '@' } : undefined,
  }));
  PRODUCT_COLUMNS.forEach((col, index) => {
    const cell = products.getRow(1).getCell(index + 1);
    cell.fill = headerFill;
    cell.font = headerFont;
    cell.alignment = { vertical: 'middle', wrapText: true };
    cell.note = col.note;
  });
  products.getRow(1).height = 32;

  catSheet.columns = [
    { header: 'Nombre *', key: 'name', width: 28 },
    { header: 'Color', key: 'color', width: 16 },
    { header: 'Icono (emoji)', key: 'icon', width: 14 },
  ];
  catSheet.getRow(1).eachCell((cell) => {
    cell.fill = headerFill;
    cell.font = headerFont;
  });
  catSheet.getCell('B1').note = 'Color del botón de la categoría en el POS.';
  for (const category of categories) {
    catSheet.addRow({
      name: category.name,
      color: category.color ? (COLOR_LABEL_BY_KEY[category.color] ?? category.color) : '',
      icon: category.icon ?? '',
    });
  }

  Object.keys(CATEGORY_COLORS).forEach((label, i) => {
    lists.getCell(i + 1, 1).value = label;
  });
  UNIT_LABELS.forEach((label, i) => {
    lists.getCell(i + 1, 2).value = label;
  });
  lists.getCell(1, 3).value = 'Sí';
  lists.getCell(2, 3).value = 'No';

  const colorCount = Object.keys(CATEGORY_COLORS).length;
  const list = (formula: string): ExcelJS.DataValidation => ({
    type: 'list',
    allowBlank: true,
    formulae: [formula],
    showErrorMessage: false,
  });
  const colIndex = (field: ProductField) => PRODUCT_COLUMNS.findIndex((c) => c.field === field) + 1;
  const letter = (field: ProductField) => products.getColumn(colIndex(field)).letter;
  const costCol = letter('cost');
  const marginCol = letter('margin');

  for (let row = 2; row <= TEMPLATE_ROWS + 1; row++) {
    const r = products.getRow(row);
    r.getCell(colIndex('category')).dataValidation = list(
      `${CATEGORIES_SHEET}!$A$2:$A$${TEMPLATE_ROWS + 1}`,
    );
    r.getCell(colIndex('track_inventory')).dataValidation = list('Listas!$C$1:$C$2');
    r.getCell(colIndex('active')).dataValidation = list('Listas!$C$1:$C$2');
    r.getCell(colIndex('unit')).dataValidation = list(`Listas!$B$1:$B$${UNIT_LABELS.length}`);
    r.getCell(colIndex('price')).value = {
      formula: `IF(OR(${costCol}${row}="",${marginCol}${row}=""),"",ROUND(${costCol}${row}*(1+${marginCol}${row}/100),-2))`,
    } as ExcelJS.CellFormulaValue;
    r.getCell(colIndex('price')).numFmt = '#,##0';
    r.getCell(colIndex('cost')).numFmt = '#,##0';
    catSheet.getRow(row).getCell(2).dataValidation = list(`Listas!$A$1:$A$${colorCount}`);
  }

  help.columns = [{ width: 110 }];
  [
    'CÓMO CARGAR PRODUCTOS',
    '',
    '1. Hoja "Categorias": escribe tus categorías y elige un color. Las que ya existen en el POS ya están listadas.',
    '2. Hoja "Productos": una fila por producto. Nombre, Categoría y Precio son obligatorios.',
    '3. Precio: escribe Costo y Margen % y el precio se calcula solo (redondeado a 100). También puedes escribir el precio encima.',
    '   Ejemplo: Costo 10.000 y Margen 50 → Precio 15.000.',
    '4. Código de barras: pon el cursor en la celda y escanea el producto con el lector.',
    '5. Stock: si lo llenas, el POS controla el inventario de ese producto.',
    '6. Guarda el archivo (.xlsx) y súbelo en Cantaro POS → Productos → Importar.',
    '',
    'Si vuelves a subir el archivo, los productos existentes se actualizan (se reconocen por código de barras, referencia o nombre + categoría). No se duplican.',
    'Las fotos se agregan después, editando cada producto.',
    '',
    'EJEMPLOS',
    'Capuchino | Café | Costo 2.500 | Margen 140 | Precio 6.000 | Unidad',
    'Vela de soya lavanda | Velas y decoración | Costo 12.000 | Margen 100 | Código 7701234567890 | Stock 10 | Stock mínimo 2',
    'Mochila tejida | Artesanías | Precio 180.000 | Stock 3',
  ].forEach((line, i) => {
    const cell = help.getCell(i + 1, 1);
    cell.value = line;
    if (line === 'CÓMO CARGAR PRODUCTOS' || line === 'EJEMPLOS') cell.font = { bold: true };
  });

  const buffer = await wb.xlsx.writeBuffer();
  return Buffer.from(buffer as ArrayBuffer);
}

// ─── Import ──────────────────────────────────────────────────────────────────

export interface ProductWorkbookResult {
  created: number;
  updated: number;
  skipped: number;
  failed: number;
  categories_created: number;
  categories_updated: number;
  errors: string[];
}

export class ProductWorkbookError extends Error {}

function findSheet(wb: ExcelJS.Workbook, name: string): ExcelJS.Worksheet | undefined {
  const target = normalizeHeader(name);
  return wb.worksheets.find((ws) => normalizeHeader(ws.name) === target);
}

export async function importProductWorkbook(
  db: Db,
  file: Buffer,
  options: { regional: RegionalSnapshot; defaultTaxCategoryId: string | null; actorUserId: string },
): Promise<ProductWorkbookResult> {
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(file as unknown as ExcelJS.Buffer);
  } catch {
    throw new ProductWorkbookError('El archivo no es un Excel (.xlsx) válido.');
  }

  const productSheet =
    findSheet(wb, PRODUCTS_SHEET) ?? wb.worksheets.find((ws) => ws.state === 'visible');
  if (!productSheet) throw new ProductWorkbookError('No se encontró la hoja "Productos".');

  const result: ProductWorkbookResult = {
    created: 0,
    updated: 0,
    skipped: 0,
    failed: 0,
    categories_created: 0,
    categories_updated: 0,
    errors: [],
  };

  type CategoryRow = { id: string; name: string; parent_id: string | null };
  const categories = db
    .prepare('SELECT id, name, parent_id FROM categories WHERE deleted_at IS NULL')
    .all() as CategoryRow[];
  let nextCategorySort =
    (
      db
        .prepare(
          'SELECT COALESCE(MAX(sort_order), 0) AS m FROM categories WHERE parent_id IS NULL AND deleted_at IS NULL',
        )
        .get() as { m: number }
    ).m + 1;

  const findCategory = (name: string): CategoryRow | 'ambiguous' | undefined => {
    const norm = name.trim().toLowerCase();
    const matches = categories.filter((c) => c.name.trim().toLowerCase() === norm);
    if (matches.length <= 1) return matches[0];
    const roots = matches.filter((c) => c.parent_id === null);
    return roots.length === 1 ? roots[0] : 'ambiguous';
  };
  const createCategory = (name: string, color: string | null, icon: string | null) => {
    const id = randomUUID();
    db.prepare(
      `INSERT INTO categories (id, name, slug, description, parent_id, color, icon, sort_order, is_active, created_at, updated_at)
       VALUES (?, ?, ?, NULL, NULL, ?, ?, ?, 1, ?, ?)`,
    ).run(id, name, slugify(name), color, icon, nextCategorySort++, now(), now());
    const row = { id, name, parent_id: null };
    categories.push(row);
    result.categories_created++;
    return row;
  };

  const findByBarcode = db.prepare(
    'SELECT id FROM products WHERE barcode = ? AND deleted_at IS NULL',
  );
  const variantBarcodeHolder = db.prepare(
    `SELECT v.product_id FROM product_variants v JOIN products p ON p.id = v.product_id
     WHERE v.barcode = ? AND v.is_active = 1 AND p.deleted_at IS NULL`,
  );
  const findBySku = db.prepare('SELECT id FROM products WHERE sku = ? AND deleted_at IS NULL');
  const findByName = db.prepare(
    'SELECT id FROM products WHERE lower(trim(name)) = lower(?) AND category_id IS ? AND deleted_at IS NULL',
  );
  const loadProduct = db.prepare(
    'SELECT id, stock_quantity, track_inventory FROM products WHERE id = ?',
  );
  const countVariants = db.prepare(
    'SELECT COUNT(*) AS n FROM product_variants WHERE product_id = ? AND is_active = 1',
  );

  const setStock = (productId: string, target: number, opening: boolean) => {
    const current = loadProduct.get(productId) as { stock_quantity: number } | undefined;
    const delta = target - Number(current?.stock_quantity ?? 0);
    if (delta === 0) return;
    adjustProductStock(db, {
      productId,
      quantityDelta: delta,
      movementType: 'adjustment',
      referenceType: opening ? 'opening_balance' : 'manual_adjustment',
      referenceId: productId,
      reason: 'Stock importado desde Excel',
      actorUserId: options.actorUserId,
    });
  };

  db.transaction(() => {
    const catSheet = findSheet(wb, CATEGORIES_SHEET);
    if (catSheet) {
      catSheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const name = cellText(row.getCell(1));
        if (!name) return;
        const colorText = normalizeHeader(cellText(row.getCell(2)));
        const color = colorText
          ? (CATEGORY_COLORS[colorText] ??
            (Object.values(CATEGORY_COLORS).includes(colorText) ? colorText : null))
          : null;
        const icon = cellText(row.getCell(3)) || null;
        const existing = findCategory(name);
        if (existing === 'ambiguous') return;
        if (!existing) {
          createCategory(name, color, icon);
          return;
        }
        const current = db
          .prepare('SELECT color, icon FROM categories WHERE id = ?')
          .get(existing.id) as { color: string | null; icon: string | null };
        const nextColor = color ?? current.color;
        const nextIcon = icon ?? current.icon;
        if (nextColor !== current.color || nextIcon !== current.icon) {
          db.prepare('UPDATE categories SET color = ?, icon = ?, updated_at = ? WHERE id = ?').run(
            nextColor,
            nextIcon,
            now(),
            existing.id,
          );
          result.categories_updated++;
        }
      });
    }

    const columns = new Map<ProductField, number>();
    productSheet.getRow(1).eachCell((cell, col) => {
      const field = PRODUCT_HEADER_ALIASES[normalizeHeader(cellText(cell))];
      if (field && !columns.has(field)) columns.set(field, col);
    });
    if (!columns.has('name') || !columns.has('price')) {
      throw new ProductWorkbookError(
        'La hoja "Productos" debe tener las columnas "Nombre" y "Precio venta". Descarga la plantilla.',
      );
    }

    const seenBarcodes = new Set<string>();
    let processed = 0;
    for (let rowNumber = 2; rowNumber <= productSheet.rowCount; rowNumber++) {
      const row = productSheet.getRow(rowNumber);
      const text = (field: ProductField) => {
        const col = columns.get(field);
        return col ? cellText(row.getCell(col)) : '';
      };
      const num = (field: ProductField) => {
        const col = columns.get(field);
        return col ? cellNumber(row.getCell(col), options.regional) : ({ empty: true } as const);
      };

      const name = text('name');
      const categoryName = text('category');
      const price = num('price');
      const cost = num('cost');
      const margin = num('margin');
      const stock = num('stock');
      const lowStock = num('low_stock');
      const barcode = text('barcode') || null;
      const sku = text('sku') || null;
      const description = text('description') || null;

      if (!name) {
        if (categoryName || !price.empty || barcode || sku) {
          result.failed++;
          result.errors.push(`Fila ${rowNumber}: falta el nombre.`);
        }
        continue;
      }
      if (++processed > MAX_IMPORT_ROWS) {
        throw new ProductWorkbookError(`El archivo supera el límite de ${MAX_IMPORT_ROWS} filas.`);
      }
      const fail = (message: string) => {
        result.failed++;
        result.errors.push(`Fila ${rowNumber} (${name}): ${message}`);
      };

      const numberError = [
        ['Costo', cost],
        ['Margen %', margin],
        ['Precio venta', price],
        ['Stock', stock],
        ['Stock mínimo', lowStock],
      ].find(
        ([, v]) =>
          !(v as ReturnType<typeof cellNumber>).empty &&
          (v as { value: number | null }).value === null,
      );
      if (numberError) {
        fail(`"${numberError[0]}" no es un número válido.`);
        continue;
      }
      const valueOf = (v: ReturnType<typeof cellNumber>) => (v.empty ? null : v.value);
      const costValue = valueOf(cost);
      const marginValue = valueOf(margin);
      let priceValue = valueOf(price);
      if (priceValue === null && costValue !== null && marginValue !== null) {
        priceValue = Math.round((costValue * (1 + marginValue / 100)) / 100) * 100;
      }
      if (priceValue === null) {
        fail('falta el precio de venta (o Costo + Margen %).');
        continue;
      }
      if (priceValue < 0 || (costValue ?? 0) < 0) {
        fail('el precio y el costo no pueden ser negativos.');
        continue;
      }
      const stockValue = valueOf(stock);
      const lowStockValue = valueOf(lowStock);
      if ((stockValue ?? 0) < 0 || (lowStockValue ?? 0) < 0) {
        fail('el stock no puede ser negativo.');
        continue;
      }

      const unitText = normalizeHeader(text('unit'));
      const saleUnit = unitText ? SALE_UNITS[unitText] : 'each';
      if (!saleUnit) {
        fail(`unidad "${text('unit')}" no válida (usa Unidad, kg, g, lb, l o ml).`);
        continue;
      }
      const trackText = text('track_inventory');
      const trackParsed = parseYesNo(trackText);
      if (trackText && trackParsed === null) {
        fail('"Controlar inventario" debe ser Sí o No.');
        continue;
      }
      const activeText = text('active');
      const activeParsed = parseYesNo(activeText);
      if (activeText && activeParsed === null) {
        fail('"Activo" debe ser Sí o No.');
        continue;
      }

      let categoryId: string | null = null;
      if (categoryName) {
        const match = findCategory(categoryName);
        if (match === 'ambiguous') {
          fail(`hay varias categorías llamadas "${categoryName}".`);
          continue;
        }
        categoryId = (match ?? createCategory(categoryName, null, null)).id;
      } else {
        fail('falta la categoría.');
        continue;
      }

      if (barcode) {
        if (seenBarcodes.has(barcode)) {
          fail(`el código de barras ${barcode} está repetido en el archivo.`);
          continue;
        }
        seenBarcodes.add(barcode);
      }

      const existingId =
        ((barcode && (findByBarcode.get(barcode) as { id: string } | undefined)?.id) ||
          (sku && (findBySku.get(sku) as { id: string } | undefined)?.id) ||
          (findByName.get(name, categoryId) as { id: string } | undefined)?.id) ??
        null;

      if (barcode) {
        const otherProduct = findByBarcode.get(barcode) as { id: string } | undefined;
        const variantHolder = variantBarcodeHolder.get(barcode) as
          | { product_id: string }
          | undefined;
        if ((otherProduct && otherProduct.id !== existingId) || variantHolder) {
          fail(`el código de barras ${barcode} ya lo usa otro producto.`);
          continue;
        }
      }

      const hasStockData = stockValue !== null || lowStockValue !== null;
      const allowFractional = saleUnit === 'each' ? 0 : 1;

      if (existingId) {
        const existing = loadProduct.get(existingId) as {
          id: string;
          stock_quantity: number;
          track_inventory: number;
        };
        const trackInventory =
          trackParsed !== null
            ? trackParsed
            : hasStockData
              ? true
              : Boolean(existing.track_inventory);
        db.prepare(
          `UPDATE products SET name = ?, category_id = ?, price = ?,
             cost = COALESCE(?, cost), sku = COALESCE(?, sku), barcode = COALESCE(?, barcode),
             description = COALESCE(?, description),
             sale_unit = CASE WHEN ? = 1 THEN ? ELSE sale_unit END,
             allow_fractional_quantity = CASE WHEN ? = 1 THEN ? ELSE allow_fractional_quantity END,
             track_inventory = ?, low_stock_threshold = COALESCE(?, low_stock_threshold),
             is_active = COALESCE(?, is_active), updated_at = ?
           WHERE id = ?`,
        ).run(
          name,
          categoryId,
          priceValue,
          costValue,
          sku,
          barcode,
          description,
          unitText ? 1 : 0,
          saleUnit,
          unitText ? 1 : 0,
          allowFractional,
          trackInventory ? 1 : 0,
          lowStockValue,
          activeParsed === null ? null : activeParsed ? 1 : 0,
          now(),
          existingId,
        );
        if (stockValue !== null) {
          if ((countVariants.get(existingId) as { n: number }).n > 0) {
            result.errors.push(
              `Fila ${rowNumber} (${name}): tiene variantes; el stock se ajusta en cada variante, no se cambió.`,
            );
          } else {
            setStock(existingId, stockValue, false);
          }
        }
        result.updated++;
        continue;
      }

      const id = generateShortId('products');
      const trackInventory = trackParsed !== null ? trackParsed : hasStockData;
      db.prepare(
        `INSERT INTO products (id, category_id, name, sku, barcode, description, price, cost,
           sale_unit, allow_fractional_quantity, weight_precision,
           tax_type, tax_rate, tax_category_id, tax_behavior, track_inventory, stock_quantity, low_stock_threshold,
           is_active, image_url, sort_order, cb_percent, tags, created_at, updated_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 3, 'none', 0, ?, 'country_default', ?, 0, ?, ?, NULL, 0, NULL, '[]', ?, ?)`,
      ).run(
        id,
        categoryId,
        name,
        sku,
        barcode,
        description,
        priceValue,
        costValue ?? 0,
        saleUnit,
        allowFractional,
        options.defaultTaxCategoryId,
        trackInventory ? 1 : 0,
        lowStockValue ?? 0,
        activeParsed === false ? 0 : 1,
        now(),
        now(),
      );
      if (stockValue) setStock(id, stockValue, true);
      result.created++;
    }
  })();

  return result;
}
