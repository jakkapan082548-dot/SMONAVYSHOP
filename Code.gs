// ====================================================================
// Google Apps Script — Backend API สำหรับระบบร้านค้าออนไลน์
// ใช้ Google Sheets เป็น Database
// ====================================================================

// ─── CONFIG ──────────────────────────────────────────────────────────
const CONFIG = {
  SHEETS: {
    PRODUCTS:  'Products',
    ORDERS:    'Orders',
    DASHBOARD: 'Dashboard',
  },
  // ค่าจัดส่งเริ่มต้น (บาท) — ปรับได้ตามต้องการ
  DEFAULT_SHIPPING_COST: 50,
  // จำนวน Order ID เริ่มต้น
  ORDER_PREFIX: 'ORD',
};

// ─── HELPERS ─────────────────────────────────────────────────────────

/**
 * สร้าง JSON Response กลับไปให้ Client
 */
function jsonResponse(data, statusCode) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

/**
 * ดึง Sheet ตามชื่อ — ถ้ายังไม่มีจะสร้างให้อัตโนมัติ
 */
function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    initSheetHeaders(sheet, name);
  }
  return sheet;
}

/**
 * ใส่ Header Row ให้ชีตที่เพิ่งสร้างใหม่
 */
function initSheetHeaders(sheet, name) {
  const headers = {
    [CONFIG.SHEETS.PRODUCTS]: [
      'ID', 'ชื่อสินค้า', 'รายละเอียด', 'ลิงก์รูปภาพ',
      'ต้นทุน', 'ราคาขาย', 'สต๊อกคงเหลือ',
    ],
    [CONFIG.SHEETS.ORDERS]: [
      'Order_ID', 'วันที่', 'รายการสินค้า', 'ยอดรวม',
      'วิธีรับของ', 'ข้อมูลลูกค้า', 'สลิปโอนเงิน', 'สถานะการตรวจสอบ',
    ],
    [CONFIG.SHEETS.DASHBOARD]: [
      'วันที่', 'ยอดขายรวม', 'ต้นทุนรวม', 'ค่าจัดส่งรวม',
      'กำไรสุทธิ', 'จำนวนออเดอร์',
    ],
  };

  if (headers[name]) {
    sheet.getRange(1, 1, 1, headers[name].length).setValues([headers[name]]);
    sheet.getRange(1, 1, 1, headers[name].length)
      .setFontWeight('bold')
      .setBackground('#4a86e8')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
}

/**
 * สร้าง Order ID ใหม่ (ORD-yyyyMMdd-XXXX)
 */
function generateOrderId() {
  const now = new Date();
  const datePart = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd');
  const sheet = getOrCreateSheet(CONFIG.SHEETS.ORDERS);
  const lastRow = sheet.getLastRow();
  const seq = String(lastRow).padStart(4, '0');
  return `${CONFIG.ORDER_PREFIX}-${datePart}-${seq}`;
}

/**
 * Sanitize string เพื่อป้องกัน injection
 */
function sanitize(value) {
  if (typeof value !== 'string') return value;
  // ตัด leading = + - @ ที่อาจถูกใช้เป็น formula injection
  return value.replace(/^[=+\-@\t\r]/, "'$&").trim();
}

// ====================================================================
// ─── doGet: ดึงข้อมูล (GET Requests) ────────────────────────────────
// ====================================================================

function doGet(e) {
  try {
    const action = (e.parameter.action || '').toLowerCase();

    switch (action) {
      case 'get_products':
        return getProducts(e);
      case 'get_product':
        return getProduct(e);
      case 'get_orders':
        return getOrders(e);
      case 'get_dashboard':
        return getDashboard(e);
      case 'refresh_dashboard':
        return refreshDashboardEndpoint(e);
      default:
        return jsonResponse({
          success: false,
          error: 'ไม่พบ action ที่ระบุ กรุณาใช้: get_products, get_product, get_orders, get_dashboard, refresh_dashboard',
        });
    }
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

/**
 * ดึงสินค้าทั้งหมด (เฉพาะที่มีสต๊อก หรือ ทั้งหมด)
 * params: ?action=get_products&in_stock=true
 */
function getProducts(e) {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.PRODUCTS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonResponse({ success: true, data: [] });
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const inStockOnly = (e.parameter.in_stock || 'false') === 'true';

  const products = data
    .map(row => ({
      id:          row[0],
      name:        row[1],
      description: row[2],
      image:       row[3],
      cost:        Number(row[4]),
      price:       Number(row[5]),
      stock:       Number(row[6]),
    }))
    .filter(p => p.id !== '' && (!inStockOnly || p.stock > 0));

  return jsonResponse({ success: true, count: products.length, data: products });
}

/**
 * ดึงสินค้าตาม ID
 * params: ?action=get_product&id=xxx
 */
function getProduct(e) {
  const id = e.parameter.id;
  if (!id) {
    return jsonResponse({ success: false, error: 'กรุณาระบุ id ของสินค้า' });
  }

  const sheet = getOrCreateSheet(CONFIG.SHEETS.PRODUCTS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonResponse({ success: false, error: 'ไม่พบสินค้า' });
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const row = data.find(r => String(r[0]) === String(id));

  if (!row) {
    return jsonResponse({ success: false, error: `ไม่พบสินค้า ID: ${id}` });
  }

  return jsonResponse({
    success: true,
    data: {
      id:          row[0],
      name:        row[1],
      description: row[2],
      image:       row[3],
      cost:        Number(row[4]),
      price:       Number(row[5]),
      stock:       Number(row[6]),
    },
  });
}

/**
 * ดึงรายการคำสั่งซื้อ (สามารถ filter ตามวันที่ หรือ สถานะ)
 * params: ?action=get_orders&status=xxx&date=yyyy-MM-dd
 */
function getOrders(e) {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.ORDERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonResponse({ success: true, data: [] });
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  let orders = data
    .filter(r => r[0] !== '')
    .map(row => ({
      orderId:     row[0],
      date:        row[1],
      items:       row[2],
      total:       Number(row[3]),
      delivery:    row[4],
      customer:    row[5],
      slip:        row[6],
      status:      row[7],
    }));

  // Filter by status
  const statusFilter = e.parameter.status;
  if (statusFilter) {
    orders = orders.filter(o => o.status === statusFilter);
  }

  // Filter by date
  const dateFilter = e.parameter.date;
  if (dateFilter) {
    orders = orders.filter(o => {
      const orderDate = Utilities.formatDate(
        new Date(o.date), Session.getScriptTimeZone(), 'yyyy-MM-dd'
      );
      return orderDate === dateFilter;
    });
  }

  return jsonResponse({ success: true, count: orders.length, data: orders });
}

/**
 * ดึงข้อมูล Dashboard (สรุปยอดรายวัน)
 * params: ?action=get_dashboard
 */
function getDashboard(e) {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.DASHBOARD);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonResponse({ success: true, data: [] });
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const rows = data
    .filter(r => r[0] !== '')
    .map(row => ({
      date:         row[0],
      totalSales:   Number(row[1]),
      totalCost:    Number(row[2]),
      totalShipping:Number(row[3]),
      netProfit:    Number(row[4]),
      orderCount:   Number(row[5]),
    }));

  return jsonResponse({ success: true, data: rows });
}

/**
 * Endpoint เพื่อสั่ง Refresh Dashboard ผ่าน GET
 */
function refreshDashboardEndpoint(e) {
  refreshDashboard();
  return jsonResponse({ success: true, message: 'Dashboard อัปเดตแล้ว' });
}

// ====================================================================
// ─── doPost: รับข้อมูล (POST Requests) ──────────────────────────────
// ====================================================================

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = (body.action || '').toLowerCase();

    switch (action) {
      case 'create_order':
        return createOrder(body);
      case 'update_order_status':
        return updateOrderStatus(body);
      case 'add_product':
        return addProduct(body);
      case 'update_stock':
        return updateStock(body);
      default:
        return jsonResponse({
          success: false,
          error: 'ไม่พบ action ที่ระบุ กรุณาใช้: create_order, update_order_status, add_product, update_stock',
        });
    }
  } catch (err) {
    return jsonResponse({ success: false, error: 'ข้อมูลไม่ถูกต้อง: ' + err.message });
  }
}

/**
 * สร้างคำสั่งซื้อใหม่
 * body: {
 *   action: "create_order",
 *   items: [{ productId: "xxx", quantity: 1 }, ...],
 *   delivery: "มารับเอง" | "จัดส่ง",
 *   customer: { name: "...", phone: "...", address: "..." },
 *   slip: "URL ของสลิปโอนเงิน"
 * }
 */
function createOrder(body) {
  // ─── Validation ───
  if (!body.items || !Array.isArray(body.items) || body.items.length === 0) {
    return jsonResponse({ success: false, error: 'กรุณาระบุรายการสินค้า (items)' });
  }
  if (!body.delivery || !['มารับเอง', 'จัดส่ง'].includes(body.delivery)) {
    return jsonResponse({ success: false, error: 'กรุณาระบุวิธีรับของ: "มารับเอง" หรือ "จัดส่ง"' });
  }
  if (!body.customer || !body.customer.name || !body.customer.phone) {
    return jsonResponse({ success: false, error: 'กรุณาระบุข้อมูลลูกค้า (name, phone)' });
  }

  // ─── ตรวจสอบสต๊อก & คำนวณยอด ───
  const productSheet = getOrCreateSheet(CONFIG.SHEETS.PRODUCTS);
  const lastRow = productSheet.getLastRow();
  if (lastRow < 2) {
    return jsonResponse({ success: false, error: 'ไม่มีสินค้าในระบบ' });
  }

  const productData = productSheet.getRange(2, 1, lastRow - 1, 7).getValues();
  let totalPrice = 0;
  let totalCost = 0;
  const itemDetails = [];

  // Lock เพื่อป้องกัน race condition ในการตัดสต๊อก
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // รอ lock สูงสุด 10 วินาที
  } catch (err) {
    return jsonResponse({ success: false, error: 'ระบบกำลังยุ่ง กรุณาลองใหม่อีกครั้ง' });
  }

  try {
    for (const item of body.items) {
      const prodIndex = productData.findIndex(r => String(r[0]) === String(item.productId));
      if (prodIndex === -1) {
        return jsonResponse({ success: false, error: `ไม่พบสินค้า ID: ${item.productId}` });
      }

      const qty = Math.max(1, Math.floor(Number(item.quantity) || 1));
      const stock = Number(productData[prodIndex][6]);
      if (stock < qty) {
        return jsonResponse({
          success: false,
          error: `สินค้า "${productData[prodIndex][1]}" เหลือสต๊อกไม่เพียงพอ (เหลือ ${stock} ชิ้น)`,
        });
      }

      const price = Number(productData[prodIndex][5]);
      const cost  = Number(productData[prodIndex][4]);

      totalPrice += price * qty;
      totalCost  += cost * qty;

      itemDetails.push({
        productId: item.productId,
        name:      productData[prodIndex][1],
        quantity:  qty,
        unitPrice: price,
        subtotal:  price * qty,
      });

      // ตัดสต๊อก
      const stockCell = productSheet.getRange(prodIndex + 2, 7); // column G
      stockCell.setValue(stock - qty);
    }

    // ─── บันทึก Order ───
    const orderSheet = getOrCreateSheet(CONFIG.SHEETS.ORDERS);
    const orderId = generateOrderId();
    const now = new Date();
    const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    const customerInfo = sanitize(
      `${body.customer.name} | ${body.customer.phone}` +
      (body.customer.address ? ` | ${body.customer.address}` : '')
    );

    const shippingCost = body.delivery === 'จัดส่ง' ? CONFIG.DEFAULT_SHIPPING_COST : 0;
    const grandTotal = totalPrice + shippingCost;

    orderSheet.appendRow([
      orderId,
      dateStr,
      sanitize(JSON.stringify(itemDetails)),
      grandTotal,
      body.delivery,
      customerInfo,
      sanitize(body.slip || ''),
      'รอตรวจสอบ',
    ]);

    // อัปเดต Dashboard หลังมี Order ใหม่
    refreshDashboard();

    return jsonResponse({
      success: true,
      message: 'สร้างคำสั่งซื้อสำเร็จ',
      data: {
        orderId:      orderId,
        items:        itemDetails,
        subtotal:     totalPrice,
        shippingCost: shippingCost,
        grandTotal:   grandTotal,
        delivery:     body.delivery,
        status:       'รอตรวจสอบ',
      },
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * อัปเดตสถานะคำสั่งซื้อ (สำหรับแอดมิน)
 * body: { action: "update_order_status", orderId: "ORD-xxx", status: "ยืนยันแล้ว" }
 */
function updateOrderStatus(body) {
  if (!body.orderId || !body.status) {
    return jsonResponse({ success: false, error: 'กรุณาระบุ orderId และ status' });
  }

  const validStatuses = ['รอตรวจสอบ', 'ยืนยันแล้ว', 'กำลังจัดส่ง', 'จัดส่งแล้ว', 'เสร็จสิ้น', 'ยกเลิก'];
  if (!validStatuses.includes(body.status)) {
    return jsonResponse({
      success: false,
      error: `สถานะไม่ถูกต้อง กรุณาใช้: ${validStatuses.join(', ')}`,
    });
  }

  const sheet = getOrCreateSheet(CONFIG.SHEETS.ORDERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonResponse({ success: false, error: 'ไม่พบคำสั่งซื้อ' });
  }

  const orderIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const rowIndex = orderIds.findIndex(id => id === body.orderId);

  if (rowIndex === -1) {
    return jsonResponse({ success: false, error: `ไม่พบคำสั่งซื้อ: ${body.orderId}` });
  }

  // อัปเดตสถานะ (column H = 8)
  sheet.getRange(rowIndex + 2, 8).setValue(body.status);

  // ถ้ายกเลิก ให้คืนสต๊อก
  if (body.status === 'ยกเลิก') {
    restoreStock(sheet.getRange(rowIndex + 2, 3).getValue());
  }

  return jsonResponse({
    success: true,
    message: `อัปเดตสถานะ ${body.orderId} เป็น "${body.status}" สำเร็จ`,
  });
}

/**
 * คืนสต๊อกเมื่อยกเลิกคำสั่งซื้อ
 */
function restoreStock(itemsJson) {
  try {
    // ลบ sanitize prefix ถ้ามี
    const cleaned = String(itemsJson).replace(/^'/, '');
    const items = JSON.parse(cleaned);
    const productSheet = getOrCreateSheet(CONFIG.SHEETS.PRODUCTS);
    const lastRow = productSheet.getLastRow();
    if (lastRow < 2) return;

    const productData = productSheet.getRange(2, 1, lastRow - 1, 7).getValues();

    for (const item of items) {
      const idx = productData.findIndex(r => String(r[0]) === String(item.productId));
      if (idx !== -1) {
        const currentStock = Number(productData[idx][6]);
        productSheet.getRange(idx + 2, 7).setValue(currentStock + item.quantity);
      }
    }
  } catch (err) {
    Logger.log('ไม่สามารถคืนสต๊อกได้: ' + err.message);
  }
}

/**
 * เพิ่มสินค้าใหม่
 * body: {
 *   action: "add_product",
 *   name: "...", description: "...", image: "URL",
 *   cost: 100, price: 250, stock: 50
 * }
 */
function addProduct(body) {
  if (!body.name || body.price === undefined || body.cost === undefined) {
    return jsonResponse({ success: false, error: 'กรุณาระบุ name, cost, price' });
  }

  const sheet = getOrCreateSheet(CONFIG.SHEETS.PRODUCTS);
  const lastRow = sheet.getLastRow();
  const newId = 'P' + String(lastRow).padStart(4, '0');

  sheet.appendRow([
    newId,
    sanitize(String(body.name)),
    sanitize(String(body.description || '')),
    sanitize(String(body.image || '')),
    Number(body.cost),
    Number(body.price),
    Number(body.stock || 0),
  ]);

  return jsonResponse({
    success: true,
    message: 'เพิ่มสินค้าสำเร็จ',
    data: { id: newId },
  });
}

/**
 * อัปเดตสต๊อกสินค้า
 * body: { action: "update_stock", productId: "P0001", stock: 100 }
 */
function updateStock(body) {
  if (!body.productId || body.stock === undefined) {
    return jsonResponse({ success: false, error: 'กรุณาระบุ productId และ stock' });
  }

  const sheet = getOrCreateSheet(CONFIG.SHEETS.PRODUCTS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonResponse({ success: false, error: 'ไม่พบสินค้า' });
  }

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const idx = ids.findIndex(id => String(id) === String(body.productId));

  if (idx === -1) {
    return jsonResponse({ success: false, error: `ไม่พบสินค้า ID: ${body.productId}` });
  }

  sheet.getRange(idx + 2, 7).setValue(Number(body.stock));

  return jsonResponse({
    success: true,
    message: `อัปเดตสต๊อก ${body.productId} เป็น ${body.stock} สำเร็จ`,
  });
}

// ====================================================================
// ─── DASHBOARD: สรุปยอดขายรายวัน ────────────────────────────────────
// ====================================================================

/**
 * คำนวณและอัปเดตข้อมูล Dashboard ใหม่ทั้งหมด
 * สรุปยอดขายรายวัน, หักต้นทุน, หักค่าจัดส่ง, แสดงกำไรสุทธิ
 */
function refreshDashboard() {
  const orderSheet = getOrCreateSheet(CONFIG.SHEETS.ORDERS);
  const dashSheet  = getOrCreateSheet(CONFIG.SHEETS.DASHBOARD);
  const productSheet = getOrCreateSheet(CONFIG.SHEETS.PRODUCTS);

  const orderLastRow = orderSheet.getLastRow();
  if (orderLastRow < 2) {
    // ล้าง Dashboard ถ้าไม่มี Order
    const dashLastRow = dashSheet.getLastRow();
    if (dashLastRow > 1) {
      dashSheet.getRange(2, 1, dashLastRow - 1, 6).clearContent();
    }
    return;
  }

  const orders = orderSheet.getRange(2, 1, orderLastRow - 1, 8).getValues();

  // สร้าง product map สำหรับหาต้นทุน
  const prodLastRow = productSheet.getLastRow();
  const productMap = {};
  if (prodLastRow >= 2) {
    const products = productSheet.getRange(2, 1, prodLastRow - 1, 7).getValues();
    products.forEach(p => {
      productMap[String(p[0])] = {
        cost:  Number(p[4]),
        price: Number(p[5]),
      };
    });
  }

  // สรุปรายวัน
  const dailySummary = {};

  for (const order of orders) {
    // ข้ามออเดอร์ที่ยกเลิก
    if (order[7] === 'ยกเลิก') continue;

    const dateKey = Utilities.formatDate(
      new Date(order[1]),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );

    if (!dailySummary[dateKey]) {
      dailySummary[dateKey] = {
        totalSales: 0,
        totalCost: 0,
        totalShipping: 0,
        orderCount: 0,
      };
    }

    const day = dailySummary[dateKey];
    day.totalSales += Number(order[3]);
    day.orderCount += 1;

    // คำนวณค่าจัดส่ง
    if (order[4] === 'จัดส่ง') {
      day.totalShipping += CONFIG.DEFAULT_SHIPPING_COST;
    }

    // คำนวณต้นทุนจากรายการสินค้า
    try {
      const cleaned = String(order[2]).replace(/^'/, '');
      const items = JSON.parse(cleaned);
      for (const item of items) {
        const prod = productMap[String(item.productId)];
        if (prod) {
          day.totalCost += prod.cost * item.quantity;
        }
      }
    } catch (err) {
      Logger.log('ไม่สามารถ parse items ของ order: ' + order[0]);
    }
  }

  // ล้าง Dashboard เดิม
  const dashLastRow = dashSheet.getLastRow();
  if (dashLastRow > 1) {
    dashSheet.getRange(2, 1, dashLastRow - 1, 6).clearContent();
  }

  // เขียนข้อมูลใหม่ (เรียงตามวันที่)
  const sortedDates = Object.keys(dailySummary).sort();
  if (sortedDates.length === 0) return;

  const rows = sortedDates.map(date => {
    const d = dailySummary[date];
    const netProfit = d.totalSales - d.totalCost - d.totalShipping;
    return [date, d.totalSales, d.totalCost, d.totalShipping, netProfit, d.orderCount];
  });

  dashSheet.getRange(2, 1, rows.length, 6).setValues(rows);

  // จัดรูปแบบตัวเลข
  if (rows.length > 0) {
    dashSheet.getRange(2, 2, rows.length, 4).setNumberFormat('#,##0.00');
    dashSheet.getRange(2, 6, rows.length, 1).setNumberFormat('#,##0');
  }
}

// ====================================================================
// ─── SETUP: ฟังก์ชันตั้งค่าเริ่มต้น ─────────────────────────────────
// ====================================================================

/**
 * เรียกใช้ฟังก์ชันนี้ครั้งแรกเพื่อสร้างชีตและ Header
 */
function setupSheets() {
  getOrCreateSheet(CONFIG.SHEETS.PRODUCTS);
  getOrCreateSheet(CONFIG.SHEETS.ORDERS);
  getOrCreateSheet(CONFIG.SHEETS.DASHBOARD);
  SpreadsheetApp.getActiveSpreadsheet().toast('สร้างชีตเรียบร้อยแล้ว!', 'Setup', 5);
}

/**
 * เพิ่มสินค้าตัวอย่างสำหรับทดสอบ
 */
function addSampleProducts() {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.PRODUCTS);
  const samples = [
    ['P0001', 'เสื้อยืดสีขาว',    'เสื้อยืดคอกลม ผ้าคอตตอน 100%',     '', 80,  250, 50],
    ['P0002', 'กางเกงยีนส์',      'กางเกงยีนส์ขายาว ทรง Slim Fit',     '', 200, 590, 30],
    ['P0003', 'รองเท้าผ้าใบ',     'รองเท้าผ้าใบสีดำ พื้นนุ่ม',          '', 350, 890, 20],
    ['P0004', 'หมวกแก๊ป',         'หมวกแก๊ปปักโลโก้ ปรับขนาดได้',      '', 50,  190, 100],
    ['P0005', 'กระเป๋าสะพายข้าง', 'กระเป๋าผ้าแคนวาส กันน้ำ',           '', 120, 390, 40],
  ];

  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, samples.length, 7).setValues(samples);
  SpreadsheetApp.getActiveSpreadsheet().toast('เพิ่มสินค้าตัวอย่างแล้ว!', 'Sample Data', 5);
}

/**
 * สร้าง Trigger สำหรับอัปเดต Dashboard อัตโนมัติทุกวัน
 */
function createDailyTrigger() {
  // ลบ trigger เก่าก่อน
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'refreshDashboard') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // สร้าง trigger ใหม่ — รันทุกวันเวลาเที่ยงคืน
  ScriptApp.newTrigger('refreshDashboard')
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'ตั้ง Trigger อัปเดต Dashboard รายวันแล้ว!', 'Trigger', 5
  );
}

// ====================================================================
// ─── CUSTOM MENU ─────────────────────────────────────────────────────
// ====================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛒 ร้านค้าออนไลน์')
    .addItem('📋 ตั้งค่าชีต', 'setupSheets')
    .addItem('📦 เพิ่มสินค้าตัวอย่าง', 'addSampleProducts')
    .addItem('📊 อัปเดต Dashboard', 'refreshDashboard')
    .addItem('⏰ ตั้ง Trigger รายวัน', 'createDailyTrigger')
    .addToUi();
}
