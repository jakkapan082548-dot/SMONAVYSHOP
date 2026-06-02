/* ====================================================
   app.js – Front‑end logic (fetch products, cart, checkout)
   ==================================================== */

/* ---------- Configuration ---------- */
// แก้ URL นี้ให้เป็น URL ที่ได้จากการ Deploy Google Apps Script ของคุณ
const API_URL = "https://script.google.com/macros/s/YOUR_DEPLOY_ID/exec";

/* ---------- Utility helpers ---------- */
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

// ฟอร์แมตค่าเป็นเงินบาท (THB)
function fmtTHB(num) {
  return Number(num).toLocaleString("th-TH", { style: "currency", currency: "THB" });
}

/* ---------- Global state ---------- */
let products = [];
let cart = [];

/* ---------- Fetch products from GAS ---------- */
async function loadProducts() {
  try {
    const res = await fetch(`${API_URL}?action=get_products`);
    const json = await res.json();
    // โครงสร้างที่ API คืน: { success:true, data:[...] }
    products = json.data || [];
    renderProducts();
  } catch (e) {
    console.error("Error loading products", e);
    alert("ไม่สามารถดึงข้อมูลสินค้าได้");
  }
}

/* ---------- Render product grid ---------- */
function renderProducts() {
  const grid = $("#productsGrid");
  grid.innerHTML = "";
  if (!products.length) {
    grid.textContent = "ไม่มีสินค้าให้แสดง";
    return;
  }
  products.forEach((p) => {
    const card = document.createElement("div");
    card.className = "product-card";
    card.innerHTML = `
      <img src="${p.image}" alt="${p.name}" loading="lazy" />
      <h3>${p.name}</h3>
      <p class="desc">${p.description || ""}</p>
      <p class="price">${fmtTHB(p.price)}</p>
      <button class="btn primary" data-id="${p.id}">เพิ่มลงตะกร้า</button>
    `;
    grid.appendChild(card);
  });

  // Bind add‑to‑cart buttons
  $$('button[data-id]').forEach((btn) => {
    btn.addEventListener("click", () => addToCart(btn.dataset.id));
  });
}

/* ---------- Cart handling ---------- */
function addToCart(id) {
  const prod = products.find((p) => p.id === id);
  if (!prod) return;
  const existing = cart.find((i) => i.id === id);
  if (existing) {
    existing.qty += 1;
  } else {
    cart.push({ id: prod.id, name: prod.name, price: prod.price, qty: 1, img: prod.image });
  }
  updateCartUI();
}

function removeFromCart(id) {
  cart = cart.filter((i) => i.id !== id);
  updateCartUI();
}

function updateCartUI() {
  const tbody = $("#cartBody");
  tbody.innerHTML = "";
  let total = 0;
  cart.forEach((item) => {
    const row = document.createElement("tr");
    total += item.price * item.qty;
    row.innerHTML = `
      <td>${item.name}</td>
      <td><input type="number" min="1" value="${item.qty}" data-id="${item.id}" class="qty-input" style="width:60px;"/></td>
      <td>${fmtTHB(item.price * item.qty)}</td>
      <td><button class="btn" data-id="${item.id}">✕</button></td>
    `;
    tbody.appendChild(row);
  });
  $("#cartTotal").textContent = fmtTHB(total);

  // Bind qty change & remove buttons
  $$(".qty-input").forEach((inp) => {
    inp.addEventListener("change", (e) => {
      const id = e.target.dataset.id;
      const qty = parseInt(e.target.value, 10);
      const item = cart.find((i) => i.id === id);
      if (item && qty > 0) item.qty = qty;
      updateCartUI();
    });
  });
  $$("button[data-id]").forEach((btn) => {
    btn.addEventListener("click", () => removeFromCart(btn.dataset.id));
  });
}

/* ---------- Navigation ---------- */
function show(sectionId) {
  $$(".section").forEach((el) => el.classList.add("hidden"));
  $(`#${sectionId}`).classList.remove("hidden");
}

/* ---------- Checkout UI ---------- */
function setupDeliveryToggle() {
  $$("input[name='deliveryMethod']").forEach((rad) => {
    rad.addEventListener("change", () => {
      const method = $("input[name='deliveryMethod']:checked").value;
      if (method === "pickup") {
        $("#pickupForm").classList.remove("hidden");
        $("#deliveryForm").classList.add("hidden");
      } else {
        $("#deliveryForm").classList.remove("hidden");
        $("#pickupForm").classList.add("hidden");
      }
    });
  });
}

function renderOrderSummary() {
  const div = $("#orderSummary");
  div.innerHTML = "";
  if (!cart.length) {
    div.textContent = "ไม่มีสินค้าในตะกร้า";
    return;
  }
  const ul = document.createElement("ul");
  cart.forEach((i) => {
    const li = document.createElement("li");
    li.textContent = `${i.name} × ${i.qty} = ${fmtTHB(i.price * i.qty)}`;
    ul.appendChild(li);
  });
  div.appendChild(ul);
}

/* ---------- Submit order ---------- */
async function submitOrder(e) {
  e.preventDefault();

  if (!cart.length) {
    alert("กรุณาเลือกสินค้าก่อน");
    return;
  }

  const method = $("input[name='deliveryMethod']:checked").value;

  // Build payload (JSON string) for items
  const items = cart.map((i) => ({ productId: i.id, quantity: i.qty }));
  const payload = {
    action: "create_order",
    items,
    delivery: method === "pickup" ? "มารับของเอง" : "จัดส่ง",
    customer: {}, // จะเติมจากฟอร์มต่อไป
  };

  // ----- เก็บข้อมูลลูกค้า -----
  if (method === "pickup") {
    payload.customer.fullName = $("#pickupForm input[name='fullName']").value.trim();
    payload.customer.studentId = $("#pickupForm input[name='studentId']").value.trim();
    payload.customer.branch   = $("#pickupForm select[name='branch']").value;
    payload.customer.phone    = $("#pickupForm input[name='phone']").value.trim();
  } else {
    payload.customer.fullName   = $("#deliveryForm input[name='fullName']").value.trim();
    payload.customer.phone      = $("#deliveryForm input[name='phone']").value.trim();
    payload.customer.address = {
      line:   $("#deliveryForm input[name='addressLine']").value.trim(),
      road:   $("#deliveryForm input[name='road']").value.trim(),
      sub:    $("#deliveryForm input[name='subdistrict']").value.trim(),
      district:$("#deliveryForm input[name='district']").value.trim(),
      province:$("#deliveryForm input[name='province']").value.trim(),
      zip:    $("#deliveryForm input[name='postalCode']").value.trim(),
    };
  }

  // ----- สลิป (ไฟล์หรือ URL) -----
  const slipFile = $("#slipFile").files[0];
  const slipUrl  = $("#slipUrl").value.trim();

  try {
    // ถ้ามีไฟล์ต้องใช้ FormData + multipart
    if (slipFile) {
      const fd = new FormData();
      fd.append('action', 'create_order');
      fd.append('payload', JSON.stringify(payload));
      fd.append('slipFile', slipFile);
      const res = await fetch(API_URL, { method: 'POST', body: fd });
      const json = await res.json();
      handleOrderResult(json);
    } else {
      // ไม่มีไฟล์，只ส่ง JSON (plain) – แต่งเป็น text/plain เพื่อให้ GAS แยกได้ง่าย
      payload.slip = slipUrl || "";
      const res = await fetch(API_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify(payload),
      });
      const json = await res.json();
      handleOrderResult(json);
    }
  } catch (err) {
    console.error(err);
    alert('การส่งออเดอร์ล้มเหลว');
  }
}

function handleOrderResult(resp) {
  if (resp.success) {
    alert(`สั่งซื้อสำเร็จ! รหัสออเดอร์: ${resp.orderId || resp.data?.orderId}`);
    cart = [];
    updateCartUI();
    show('product-list');
  } else {
    alert('เกิดข้อผิดพลาด: ' + (resp.error || resp.message || 'ไม่ทราบสาเหตุ'));
  }
}

/* ---------- Initialize ---------- */
function init() {
  loadProducts();
  show('product-list');

  // ปุ่มนำทาง
  $("#proceedBtn").addEventListener('click', () => {
    renderOrderSummary();
    show('checkout-section');
    setupDeliveryToggle(); // แสดงฟอร์มตามค่าเริ่มต้น
  });
  $("#backToProductsBtn").addEventListener('click', () => show('product-list'));
  $("#backToCartBtn").addEventListener('click', () => show('cart-section'));

  // ฟอร์ม checkout submit
  $("#orderForm").addEventListener('submit', submitOrder);
}

document.addEventListener('DOMContentLoaded', init);
