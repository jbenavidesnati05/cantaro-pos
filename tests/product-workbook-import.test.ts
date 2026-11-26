/** Excel (.xlsx) product template and import. */
const Module = require('module');
const originalLoad = Module._load;
const fs = require('fs');
const os = require('os');
const path = require('path');
const testDir = fs.mkdtempSync(path.join(os.tmpdir(), 'flo-product-workbook-'));

Module._load = function (request: string, parent: unknown, isMain: boolean) {
  if (request === 'electron') {
    return { app: { isPackaged: true, getPath: () => testDir, getVersion: () => 'test' } };
  }
  return originalLoad.apply(this, arguments as any);
};

const ExcelJS = require('exceljs');
const {
  initTestDb,
  createApp,
  startServer,
  seedOwnerUser,
  api,
  assertEqualOrThrow,
  assertOrThrow,
  closeDatabase,
  now,
} = require('./helpers/test-setup');
const { menuCsvRoutes } = require('../main/routes/menu-csv');
const { validateInventoryLedgerDatabase } = require('../main/db');

async function workbookBase64(fill: (products: any, categories: any) => void): Promise<string> {
  const wb = new ExcelJS.Workbook();
  const products = wb.addWorksheet('Productos');
  const categories = wb.addWorksheet('Categorias');
  products.addRow([
    'Nombre *',
    'Categoría *',
    'Costo',
    'Margen %',
    'Precio venta *',
    'Código de barras',
    'Referencia (SKU)',
    'Controlar inventario',
    'Stock',
    'Stock mínimo',
    'Unidad',
    'Descripción',
    'Activo',
  ]);
  categories.addRow(['Nombre *', 'Color', 'Icono (emoji)']);
  fill(products, categories);
  return Buffer.from(await wb.xlsx.writeBuffer()).toString('base64');
}

async function main() {
  console.log('Integration Test: Excel product template and import');

  const db = initTestDb();
  const setSetting = (key: string, value: string) =>
    db
      .prepare('INSERT OR REPLACE INTO settings (key, value, updated_at) VALUES (?, ?, ?)')
      .run(key, value, now());
  setSetting('country', 'CO');
  setSetting('currency', 'COP');
  setSetting('timezone', 'America/Bogota');
  const { authHeader } = seedOwnerUser(db);

  const app = createApp({ '/api/menu-csv': menuCsvRoutes });
  const { baseUrl, server } = await startServer(app);

  try {
    const templateRes = await (globalThis as any).fetch(
      `${baseUrl}/api/menu-csv/template-xlsx/products`,
      { headers: authHeader },
    );
    assertEqualOrThrow(templateRes.status, 200, 'template downloads');
    const template = new ExcelJS.Workbook();
    await template.xlsx.load(Buffer.from(await templateRes.arrayBuffer()));
    const sheet = template.getWorksheet('Productos');
    assertOrThrow(Boolean(sheet), 'template has a Productos sheet');
    assertOrThrow(Boolean(template.getWorksheet('Categorias')), 'template has a Categorias sheet');
    assertEqualOrThrow(sheet.getCell('A1').value, 'Nombre *', 'first header is the name');
    assertEqualOrThrow(sheet.getCell('B2').dataValidation?.type, 'list', 'category is a dropdown');

    const file = await workbookBase64((products, categories) => {
      categories.addRow(['Café', 'naranja', '☕']);
      products.addRow([
        'Capuchino',
        'Café',
        2500,
        140,
        { formula: 'IF(OR(C2="",D2=""),"",ROUND(C2*(1+D2/100),-2))' },
      ]);
      products.addRow([
        'Vela lavanda',
        'Velas',
        12000,
        100,
        null,
        7701234567890,
        null,
        null,
        10,
        2,
      ]);
      products.addRow(['Mochila', 'Artesanías', null, null, '180.000', null, 'MOC-1', null, 3]);
      products.addRow(['Sin precio', 'Café']);
    });

    const first = await api(baseUrl, '/api/menu-csv/import-xlsx/products', {
      method: 'POST',
      headers: authHeader,
      body: { file },
    });
    assertEqualOrThrow(first.status, 200, 'import succeeds');
    assertEqualOrThrow(first.data.created, 3, 'three products created');
    assertEqualOrThrow(first.data.failed, 1, 'row without a price fails');
    assertEqualOrThrow(first.data.categories_created, 3, 'missing categories are created');

    const byName = (name: string) => db.prepare('SELECT * FROM products WHERE name = ?').get(name);
    assertEqualOrThrow(byName('Capuchino').price, 6000, 'price comes from cost and margin');
    const vela = byName('Vela lavanda');
    assertEqualOrThrow(vela.price, 24000, 'empty price is computed server side');
    assertEqualOrThrow(vela.barcode, '7701234567890', 'numeric barcode keeps every digit');
    assertEqualOrThrow(vela.track_inventory, 1, 'stock turns on inventory tracking');
    assertEqualOrThrow(vela.stock_quantity, 10, 'opening stock recorded');
    assertEqualOrThrow(vela.low_stock_threshold, 2, 'low stock threshold recorded');
    assertEqualOrThrow(byName('Mochila').price, 180000, 'Colombian thousands separator parsed');
    assertEqualOrThrow(
      db.prepare("SELECT color FROM categories WHERE name = 'Café'").get().color,
      'orange',
      'Spanish color name maps to the palette key',
    );

    const second = await api(baseUrl, '/api/menu-csv/import-xlsx/products', {
      method: 'POST',
      headers: authHeader,
      body: {
        file: await workbookBase64((products) => {
          products.addRow([
            'Vela lavanda',
            'Velas',
            12000,
            null,
            25000,
            '7701234567890',
            null,
            null,
            7,
          ]);
        }),
      },
    });
    assertEqualOrThrow(second.data.updated, 1, 're-import updates by barcode');
    assertEqualOrThrow(second.data.created, 0, 're-import does not duplicate');
    const updated = byName('Vela lavanda');
    assertEqualOrThrow(updated.price, 25000, 'price updated');
    assertEqualOrThrow(updated.stock_quantity, 7, 'stock adjusted through the ledger');
    assertEqualOrThrow(updated.low_stock_threshold, 2, 'blank cell keeps the stored value');
    assertEqualOrThrow(
      db.prepare("SELECT COUNT(*) AS n FROM products WHERE name = 'Vela lavanda'").get().n,
      1,
      'still one product',
    );
    assertEqualOrThrow(
      validateInventoryLedgerDatabase(db),
      null,
      'inventory ledger stays consistent',
    );

    const bad = await api(baseUrl, '/api/menu-csv/import-xlsx/products', {
      method: 'POST',
      headers: authHeader,
      body: { file: Buffer.from('not a workbook').toString('base64') },
    });
    assertEqualOrThrow(bad.status, 400, 'a non-xlsx file is rejected');
  } finally {
    server.close();
    closeDatabase();
  }
  console.log('\nAll Excel product import checks passed');
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
