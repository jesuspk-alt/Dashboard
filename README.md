<!DOCTYPE html>
<html lang="th">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Stock Inventory Dashboard</title>


  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js"></script>

  <style>
    :root {
      --navy: #132238;
      --blue: #2563eb;
      --bg: #f4f7fb;
      --white: #ffffff;
      --text: #1f2937;
      --muted: #6b7280;
      --border: #e5e7eb;
      --green: #16a34a;
      --orange: #f59e0b;
      --red: #dc2626;
      --purple: #7c3aed;
    }

    * {
      box-sizing: border-box;
    }

    body {
      margin: 0;
      font-family: Arial, "Noto Sans Thai", sans-serif;
      background: var(--bg);
      color: var(--text);
    }

    .header {
      background: var(--navy);
      color: white;
      padding: 22px 32px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
    }

    .header h1 {
      margin: 0;
      font-size: 24px;
    }

    .header p {
      margin: 6px 0 0;
      color: #cbd5e1;
      font-size: 14px;
    }

    .updated {
      font-size: 13px;
      color: #cbd5e1;
      white-space: nowrap;
    }

    .container {
      max-width: 1500px;
      margin: 0 auto;
      padding: 24px 32px 40px;
    }

    .filters {
      background: var(--white);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 16px;
      display: flex;
      flex-wrap: wrap;
      gap: 12px;
      margin-bottom: 20px;
      align-items: end;
    }

    .filter-group {
      display: flex;
      flex-direction: column;
      gap: 6px;
    }

    .filter-group label {
      font-size: 12px;
      color: var(--muted);
      font-weight: bold;
    }

    select,
    input {
      min-width: 180px;
      padding: 10px 12px;
      border: 1px solid var(--border);
      border-radius: 7px;
      background: white;
      color: var(--text);
      font-size: 14px;
    }

    button {
      padding: 10px 16px;
      border: 0;
      border-radius: 7px;
      background: var(--blue);
      color: white;
      cursor: pointer;
      font-weight: bold;
      font-size: 14px;
    }

    button:hover {
      background: #1d4ed8;
    }

    .btn-secondary {
      background: #64748b;
    }

    .btn-secondary:hover {
      background: #475569;
    }

    .btn-danger {
      background: var(--red);
    }

    .btn-danger:hover {
      background: #b91c1c;
    }

    .btn-edit {
      background: #f59e0b;
      padding: 7px 10px;
      font-size: 12px;
    }

    .btn-edit:hover {
      background: #d97706;
    }

    .btn-delete {
      background: #dc2626;
      padding: 7px 10px;
      font-size: 12px;
    }

    .btn-delete:hover {
      background: #b91c1c;
    }
.btn-success {
  background: #16a34a;
}

.btn-success:hover {
  background: #15803d;
}

.btn-warning {
  background: #f59e0b;
}

.btn-warning:hover {
  background: #d97706;
}

.transaction-in {
  color: #166534;
  font-weight: bold;
}

.transaction-out {
  color: #b45309;
  font-weight: bold;
}

    .kpi-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(190px, 1fr));
      gap: 16px;
      margin-bottom: 20px;
    }

    .kpi-card {
      background: var(--white);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 18px;
      border-left: 5px solid var(--blue);
    }

    .kpi-card.red {
      border-left-color: var(--red);
    }

    .kpi-card.orange {
      border-left-color: var(--orange);
    }

    .kpi-card.green {
      border-left-color: var(--green);
    }

    .kpi-card.purple {
      border-left-color: var(--purple);
    }

    .kpi-label {
      color: var(--muted);
      font-size: 13px;
      margin-bottom: 8px;
    }

    .kpi-value {
      font-size: 25px;
      font-weight: bold;
    }

    .kpi-sub {
      margin-top: 6px;
      font-size: 12px;
      color: var(--muted);
    }

    .grid-2 {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 20px;
      margin-bottom: 20px;
    }

    .panel {
      background: var(--white);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 20px;
    }

    .panel-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
      margin-bottom: 16px;
    }

    .panel h2 {
      margin: 0;
      font-size: 17px;
    }

    .chart-wrap {
      height: 300px;
    }

    .table-wrap {
      overflow-x: auto;
    }

    table {
      border-collapse: collapse;
      width: 100%;
      font-size: 14px;
    }

    th {
      background: #f8fafc;
      color: #475569;
      text-align: left;
      font-size: 12px;
      white-space: nowrap;
    }

    th,
    td {
      padding: 12px;
      border-bottom: 1px solid var(--border);
    }

    td.number,
    th.number {
      text-align: right;
    }

    tr:hover td {
      background: #fafcff;
    }

    .badge {
      display: inline-block;
      padding: 5px 9px;
      border-radius: 999px;
      font-size: 12px;
      font-weight: bold;
      white-space: nowrap;
    }

    .badge-normal {
      color: #166534;
      background: #dcfce7;
    }

    .badge-low {
      color: #92400e;
      background: #fef3c7;
    }

    .badge-out {
      color: #991b1b;
      background: #fee2e2;
    }

    .badge-over {
      color: #5b21b6;
      background: #ede9fe;
    }

    .action-buttons {
      display: flex;
      gap: 6px;
    }

    .empty {
      text-align: center;
      color: var(--muted);
      padding: 28px;
    }

    /* Modal */
    .modal {
      display: none;
      position: fixed;
      z-index: 1000;
      inset: 0;
      background: rgba(15, 23, 42, 0.55);
      padding: 20px;
      overflow-y: auto;
    }

    .modal.show {
      display: flex;
      align-items: center;
      justify-content: center;
    }

    .modal-content {
      width: min(850px, 100%);
      background: white;
      border-radius: 14px;
      padding: 24px;
      box-shadow: 0 20px 50px rgba(0, 0, 0, 0.2);
    }

    .modal-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 20px;
    }

    .modal-header h2 {
      margin: 0;
      font-size: 20px;
    }

    .close-btn {
      background: transparent;
      color: #64748b;
      padding: 0;
      font-size: 28px;
      line-height: 1;
    }

    .close-btn:hover {
      background: transparent;
      color: #111827;
    }

    .form-grid {
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 16px;
    }

    .form-group {
      display: flex;
      flex-direction: column;
      gap: 6px;
    }

    .form-group label {
      font-size: 13px;
      font-weight: bold;
      color: #475569;
    }

    .form-group input {
      width: 100%;
      min-width: 0;
    }

    .form-group.full-width {
      grid-column: span 3;
    }

    .form-actions {
      display: flex;
      justify-content: flex-end;
      gap: 10px;
      margin-top: 24px;
      padding-top: 18px;
      border-top: 1px solid var(--border);
    }

    @media (max-width: 900px) {
      .header {
        padding: 18px;
        display: block;
      }

      .updated {
        margin-top: 10px;
      }

      .container {
        padding: 18px;
      }

      .grid-2 {
        grid-template-columns: 1fr;
      }

      .form-grid {
        grid-template-columns: 1fr;
      }

      .form-group.full-width {
        grid-column: span 1;
      }
    }
  </style>
</head>

<body>
  <header class="header">
    <div>
      <h1>📦 Stock Inventory Dashboard</h1>
      <p>ภาพรวมสต๊อกสินค้า สถานะสินค้า และรายการที่ต้องดำเนินการ</p>
    </div>
    <div class="updated" id="updatedDate"></div>
  </header>

  <main class="container">
    <section class="filters">
      <div class="filter-group">
        <label for="warehouseFilter">คลังสินค้า</label>
        <select id="warehouseFilter"></select>
      </div>

      <div class="filter-group">
        <label for="categoryFilter">หมวดหมู่</label>
        <select id="categoryFilter"></select>
      </div>

      <div class="filter-group">
        <label for="statusFilter">สถานะสต๊อก</label>
        <select id="statusFilter">
          <option value="ทั้งหมด">ทั้งหมด</option>
          <option value="ปกติ">ปกติ</option>
          <option value="ควรสั่งซื้อ">ควรสั่งซื้อ</option>
          <option value="หมดสต๊อก">หมดสต๊อก</option>
          <option value="สต๊อกเกิน">สต๊อกเกิน</option>
        </select>
      </div>

      <button onclick="resetFilters()" class="btn-secondary">ล้างตัวกรอง</button>
      <button onclick="openReceiveModal()" class="btn-success">+ รับเข้าสินค้า</button>
      <button onclick="openIssueModal()" class="btn-warning">- เบิกสินค้า</button>
<button type="button" class="btn-secondary" onclick="document.getElementById('importExcelFile').click()">
   Import Excel
   </button>
<input

    type="file"
        id="importExcelFile"
        accept=".xlsx,.xls,.csv"
        style="display: none;"
        onchange="importStockFromExcel(event)"
       />
      <button onclick="openAddModal()">+ เพิ่มข้อมูลสินค้า</button>
    </section>

    <section class="kpi-grid">
      <div class="kpi-card">
        <div class="kpi-label">มูลค่าสต๊อกรวม</div>
        <div class="kpi-value" id="totalStockValue">฿0</div>
        <div class="kpi-sub">คำนวณจากต้นทุนสินค้า</div>
      </div>

      <div class="kpi-card green">
        <div class="kpi-label">จำนวน SKU ทั้งหมด</div>
        <div class="kpi-value" id="totalSku">0</div>
        <div class="kpi-sub">ตามเงื่อนไขที่เลือก</div>
      </div>

      <div class="kpi-card red">
        <div class="kpi-label">สินค้าหมดสต๊อก</div>
        <div class="kpi-value" id="outOfStock">0</div>
        <div class="kpi-sub">ต้องตรวจสอบและเร่งสั่งซื้อ</div>
      </div>

      <div class="kpi-card orange">
        <div class="kpi-label">สินค้าที่ควรสั่งซื้อ</div>
        <div class="kpi-value" id="lowStock">0</div>
        <div class="kpi-sub">ต่ำกว่าจุดสั่งซื้อ</div>
      </div>

      <div class="kpi-card purple">
        <div class="kpi-label">มูลค่าสต๊อกเกิน</div>
        <div class="kpi-value" id="overStockValue">฿0</div>
        <div class="kpi-sub">สินค้าเกินระดับ Max Stock</div>
      </div>
    </section>

    <section class="grid-2">
      <div class="panel">
        <h2>มูลค่าสต๊อกตามหมวดสินค้า</h2>
        <div class="chart-wrap">
          <canvas id="categoryChart"></canvas>
        </div>
      </div>

      <div class="panel">
        <h2>สถานะสินค้า</h2>
        <div class="chart-wrap">
          <canvas id="statusChart"></canvas>
        </div>
      </div>
    </section>

    <section class="panel">
      <div class="panel-header">
        <h2>รายการสินค้าที่ต้องติดตาม</h2>
        <button onclick="openAddModal()">+ เพิ่มข้อมูลสินค้า</button>
      </div>
<div style="display: flex; gap: 8px;">

    <button

      type="button"

      class="btn-success"

      onclick="exportStockExcel()"

    >

      Export Excel

    </button>

  </div>

</div>

      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>SKU</th>
              <th>สินค้า</th>
              <th>หมวดหมู่</th>
              <th>คลัง</th>
              <th class="number">คงเหลือ</th>
              <th class="number">น้ำหนัก (กก.)</th>
              <th class="number">จุดสั่งซื้อ</th>
              <th class="number">แนะนำสั่ง</th>
              <th class="number">มูลค่าสต๊อก</th>
              <th>สถานะ</th>
              <th>จัดการ</th>
            </tr>
          </thead>
          <tbody id="stockTableBody"></tbody>
        </table>
      </div>
    </section>
<section class="panel" style="margin-top: 20px;">
  <div class="panel-header">
    <h2>ประวัติการรับเข้าและเบิกสินค้า</h2>

    <!-- ปุ่ม Export อยู่ด้านขวาของหัวข้อ -->
    <div style="display: flex; gap: 8px;">
      <button
        type="button"
        onclick="exportTransactionsExcel()"
        class="btn-success"
      >
        Export Excel
      </button>

      <button
        type="button"
        onclick="exportTransactionsCSV()"
        class="btn-secondary"
      >
        Export CSV
      </button>
    </div>
  </div>

  <div class="table-wrap">
    <table>
      <thead>
        <tr>
          <th>วันที่</th>
          <th>ประเภท</th>
          <th>เลขที่เอกสาร</th>
          <th>SKU</th>
          <th>สินค้า</th>
          <th>คลังสินค้า</th>
          <th class="number">จำนวน</th>
          <th>หมายเหตุ</th>
        </tr>
      </thead>

      <tbody id="transactionTableBody">
        <!-- JavaScript จะแสดงข้อมูลประวัติตรงนี้ -->
      </tbody>
    </table>
  </div>
</section>
  </main>

  <div class="modal" id="productModal">
    <div class="modal-content">
      <div class="modal-header">
        <h2 id="modalTitle">เพิ่มข้อมูลสินค้า</h2>
        <button class="close-btn" onclick="closeModal()" type="button">&times;</button>
      </div>

      <form id="productForm">
        <input type="hidden" id="editIndex" />

        <div class="form-grid">
          <div class="form-group">
            <label for="sku">SKU *</label>
            <input type="text" id="sku" required placeholder="เช่น SKU-007" />
          </div>

          <div class="form-group full-width">
            <label for="productName">ชื่อสินค้า *</label>
            <input type="text" id="productName" required placeholder="ระบุชื่อสินค้า" />
          </div>

          <div class="form-group">
            <label for="category">หมวดหมู่ *</label>
            <input type="text" id="category" required placeholder="เช่น เครื่องดื่ม" />
          </div>

          <div class="form-group">
            <label for="warehouse">คลังสินค้า *</label>
            <input type="text" id="warehouse" required placeholder="เช่น คลังหลัก" />
          </div>

          <div class="form-group">
            <label for="unitCost">ต้นทุนต่อหน่วย (บาท) *</label>
            <input type="number" id="unitCost" min="0" step="0.01" required />
          </div>

          <div class="form-group">
            <label for="stockOnHand">จำนวนคงเหลือ *</label>
            <input type="number" id="stockOnHand" min="0" step="1" required />
          </div>
<div class="form-group">
  <label for="weightKg">น้ำหนัก (กก.)</label>
  <input
    type="number"
    id="weightKg"
    min="0"
    step="0.01"
    value="0"
    placeholder="เช่น 25.50"
  />
</div>

          <div class="form-group">
            <label for="reservedQty">จำนวนถูกจอง</label>
            <input type="number" id="reservedQty" min="0" step="1" value="0" />
          </div>

          <div class="form-group">
            <label for="inTransitQty">ของระหว่างทาง</label>
            <input type="number" id="inTransitQty" min="0" step="1" value="0" />
          </div>

          <div class="form-group">
            <label for="reorderPoint">จุดสั่งซื้อ (Reorder Point) *</label>
            <input type="number" id="reorderPoint" min="0" step="1" required />
          </div>

          <div class="form-group">
            <label for="maxStock">สต๊อกสูงสุด (Max Stock) *</label>
            <input type="number" id="maxStock" min="0" step="1" required />
          </div>

          <div class="form-group">
            <label for="avgDailySales">ยอดขายเฉลี่ยต่อวัน *</label>
            <input type="number" id="avgDailySales" min="0" step="0.01" required />
          </div>
        </div>

        <div class="form-actions">
          <button type="button" class="btn-secondary" onclick="closeModal()">ยกเลิก</button>
          <button type="submit">บันทึกข้อมูล</button>
        </div>
      </form>
    </div>
  </div>
<!-- Modal รับเข้า / เบิกสินค้า -->
<div class="modal" id="transactionModal">
  <div class="modal-content">
    <div class="modal-header">
      <h2 id="transactionModalTitle">รับเข้าสินค้า</h2>
      <button class="close-btn" onclick="closeTransactionModal()" type="button">&times;</button>
    </div>

    <form id="transactionForm">
      <input type="hidden" id="transactionType" />

      <div class="form-grid">
        <div class="form-group full-width">
          <label for="transactionSku">สินค้า *</label>
          <select id="transactionSku" required></select>
        </div>

        <div class="form-group">
          <label for="transactionQty">จำนวน *</label>
          <input type="number" id="transactionQty" min="1" step="1" required placeholder="ระบุจำนวน" />
        </div>

        <div class="form-group">
          <label for="transactionDate">วันที่ทำรายการ *</label>
          <input type="date" id="transactionDate" required />
        </div>

        <div class="form-group">
          <label for="transactionDocNo">เลขที่เอกสาร</label>
          <input type="text" id="transactionDocNo" placeholder="เช่น GR-001 หรือ IV-001" />
        </div>

        <div class="form-group full-width">
          <label for="transactionRemark">หมายเหตุ</label>
          <input type="text" id="transactionRemark" placeholder="เช่น รับจาก Supplier A / เบิกใช้สำหรับสาขา A" />
        </div>
      </div>

      <div class="form-actions">
        <button type="button" class="btn-secondary" onclick="closeTransactionModal()">ยกเลิก</button>
        <button type="submit" id="transactionSubmitButton">บันทึกรับสินค้า</button>
      </div>
    </form>
  </div>
</div>
  <script>
    const STORAGE_KEY = "stockDashboardData";
        const TRANSACTION_STORAGE_KEY = "stockDashboardTransactions";

    const defaultStockData = [
      {
        sku: "SKU-001",
        productName: "น้ำดื่ม 600 ml",
        category: "เครื่องดื่ม",
        warehouse: "คลังหลัก",
        stockOnHand: 25,
weightKg: 0,
        reservedQty: 3,
        inTransitQty: 50,
        unitCost: 8,
        reorderPoint: 30,
        maxStock: 150,
        avgDailySales: 12
      },
      {
        sku: "SKU-002",
        productName: "มันฝรั่งทอดรสคลาสสิก",
        category: "ขนม",
        warehouse: "คลังหลัก",
        stockOnHand: 10,
weightKg: 0,
        reservedQty: 0,
        inTransitQty: 0,
        unitCost: 15,
        reorderPoint: 20,
        maxStock: 80,
        avgDailySales: 3
      },
      {
        sku: "SKU-003",
        productName: "กาแฟกระป๋อง",
        category: "เครื่องดื่ม",
        warehouse: "คลังหลัก",
        stockOnHand: 0,
weightKg: 0,
        reservedQty: 0,
        inTransitQty: 100,
        unitCost: 18,
        reorderPoint: 25,
        maxStock: 120,
        avgDailySales: 5
      },
      {
        sku: "SKU-004",
        productName: "สบู่เหลว 500 ml",
        category: "ของใช้ในบ้าน",
        warehouse: "คลังสาขา A",
        stockOnHand: 150,
        reservedQty: 5,
weightKg: 0,
        inTransitQty: 0,
        unitCost: 45,
        reorderPoint: 20,
        maxStock: 100,
        avgDailySales: 2
      }
    ];

    let stockData = JSON.parse(localStorage.getItem(STORAGE_KEY)) || defaultStockData;
        let transactions = JSON.parse(localStorage.getItem(TRANSACTION_STORAGE_KEY)) || [];
    let categoryChart;
    let statusChart;

    const moneyFormat = new Intl.NumberFormat("th-TH", {
      style: "currency",
      currency: "THB",
      minimumFractionDigits: 0,
      maximumFractionDigits: 0
    });

    const numberFormat = new Intl.NumberFormat("th-TH", {
      maximumFractionDigits: 1
    });

    function saveData() {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(stockData));
    }

    function calculateItem(item) {
      const availableQty = item.stockOnHand - item.reservedQty;
      const stockValue = item.stockOnHand * item.unitCost;
      const daysOfCover = item.avgDailySales > 0
        ? availableQty / item.avgDailySales
        : 0;

      let status = "ปกติ";

      if (availableQty <= 0) {
        status = "หมดสต๊อก";
      } else if (availableQty <= item.reorderPoint) {
        status = "ควรสั่งซื้อ";
      } else if (availableQty > item.maxStock) {
        status = "สต๊อกเกิน";
      }

      const suggestedOrderQty = Math.max(
        0,
        item.maxStock - (availableQty + item.inTransitQty)
      );

      return {
        ...item,
        availableQty,
        stockValue,
        daysOfCover,
        status,
        suggestedOrderQty
      };
    }

    function updateFilterOptions() {
      const warehouseFilter = document.getElementById("warehouseFilter");
      const categoryFilter = document.getElementById("categoryFilter");

      const selectedWarehouse = warehouseFilter.value || "ทั้งหมด";
      const selectedCategory = categoryFilter.value || "ทั้งหมด";

      const warehouses = [...new Set(stockData.map(item => item.warehouse))].sort();
      const categories = [...new Set(stockData.map(item => item.category))].sort();

      warehouseFilter.innerHTML =
        `<option value="ทั้งหมด">ทั้งหมด</option>` +
        warehouses.map(value => `<option value="${value}">${value}</option>`).join("");

      categoryFilter.innerHTML =
        `<option value="ทั้งหมด">ทั้งหมด</option>` +
        categories.map(value => `<option value="${value}">${value}</option>`).join("");

      warehouseFilter.value = warehouses.includes(selectedWarehouse)
        ? selectedWarehouse
        : "ทั้งหมด";

      categoryFilter.value = categories.includes(selectedCategory)
        ? selectedCategory
        : "ทั้งหมด";
    }

    function getFilteredData() {
      const warehouse = document.getElementById("warehouseFilter").value;
      const category = document.getElementById("categoryFilter").value;
      const status = document.getElementById("statusFilter").value;

      return stockData
        .map((item, index) => ({ ...calculateItem(item), index }))
        .filter(item =>
          (warehouse === "ทั้งหมด" || item.warehouse === warehouse) &&
          (category === "ทั้งหมด" || item.category === category) &&
          (status === "ทั้งหมด" || item.status === status)
        );
    }

    function getBadge(status) {
      const badgeClass = {
        "ปกติ": "badge-normal",
        "ควรสั่งซื้อ": "badge-low",
        "หมดสต๊อก": "badge-out",
        "สต๊อกเกิน": "badge-over"
      }[status];

      return `<span class="badge ${badgeClass}">${status}</span>`;
    }

    function renderKpis(data) {
      const totalStockValue = data.reduce((sum, item) => sum + item.stockValue, 0);
      const outOfStock = data.filter(item => item.status === "หมดสต๊อก").length;
      const lowStock = data.filter(item => item.status === "ควรสั่งซื้อ").length;
      const overStockValue = data
        .filter(item => item.status === "สต๊อกเกิน")
        .reduce((sum, item) => sum + item.stockValue, 0);

      document.getElementById("totalStockValue").textContent = moneyFormat.format(totalStockValue);
      document.getElementById("totalSku").textContent = numberFormat.format(data.length);
      document.getElementById("outOfStock").textContent = numberFormat.format(outOfStock);
      document.getElementById("lowStock").textContent = numberFormat.format(lowStock);
      document.getElementById("overStockValue").textContent = moneyFormat.format(overStockValue);
    }

    function renderTable(data) {
      const tbody = document.getElementById("stockTableBody");

      const priority = {
        "หมดสต๊อก": 1,
        "ควรสั่งซื้อ": 2,
        "สต๊อกเกิน": 3,
        "ปกติ": 4
      };

      const sortedData = [...data].sort((a, b) => {
  return a.sku.localeCompare(b.sku, undefined, {
    numeric: true,
    sensitivity: "base"
  });
});


      if (sortedData.length === 0) {
        tbody.innerHTML = `
          <tr>
            <td colspan="11" class="empty">ไม่พบข้อมูลตามเงื่อนไขที่เลือก</td>
          </tr>
        `;
        return;
      }

      tbody.innerHTML = sortedData.map(item => `
        <tr>
          <td>${item.sku}</td>
          <td><strong>${item.productName}</strong></td>
          <td>${item.category}</td>
          <td>${item.warehouse}</td>
          <td class="number">${numberFormat.format(item.availableQty)}</td>
          <td class="number">${numberFormat.format(item.weightKg || 0)} กก.</td>
          <td class="number">${numberFormat.format(item.reorderPoint)}</td>
          <td class="number">${numberFormat.format(item.suggestedOrderQty)}</td>
          <td class="number">${moneyFormat.format(item.stockValue)}</td>
          <td>${getBadge(item.status)}</td>
          <td>
            <div class="action-buttons">
              <button class="btn-edit" onclick="openEditModal(${item.index})">แก้ไข</button>
              <button class="btn-delete" onclick="deleteItem(${item.index})">ลบ</button>
            </div>
          </td>
        </tr>
      `).join("");
    }

    function renderCharts(data) {
      const categorySummary = {};

      data.forEach(item => {
        categorySummary[item.category] = (categorySummary[item.category] || 0) + item.stockValue;
      });

      const statusSummary = {
        "ปกติ": 0,
        "ควรสั่งซื้อ": 0,
        "หมดสต๊อก": 0,
        "สต๊อกเกิน": 0
      };

      data.forEach(item => {
        statusSummary[item.status]++;
      });

      if (categoryChart) categoryChart.destroy();
      if (statusChart) statusChart.destroy();

      categoryChart = new Chart(document.getElementById("categoryChart"), {
        type: "bar",
        data: {
          labels: Object.keys(categorySummary),
          datasets: [{
            label: "มูลค่าสต๊อก (บาท)",
            data: Object.values(categorySummary),
            backgroundColor: ["#2563eb", "#16a34a", "#f59e0b", "#7c3aed", "#dc2626"],
            borderRadius: 6
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                label: context => moneyFormat.format(context.raw)
              }
            }
          },
          scales: {
            y: {
              beginAtZero: true,
              ticks: {
                callback: value => "฿" + numberFormat.format(value)
              }
            }
          }
        }
      });

      statusChart = new Chart(document.getElementById("statusChart"), {
        type: "doughnut",
        data: {
          labels: Object.keys(statusSummary),
          datasets: [{
            data: Object.values(statusSummary),
            backgroundColor: ["#16a34a", "#f59e0b", "#dc2626", "#7c3aed"],
            borderWidth: 0
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { position: "bottom" }
          }
        }
      });
    }

    function updateDashboard() {
      updateFilterOptions();

      const data = getFilteredData();
      renderKpis(data);
      renderTable(data);
      renderCharts(data);
            renderTransactionTable();

      document.getElementById("updatedDate").textContent =
        "อัปเดตล่าสุด: " + new Date().toLocaleString("th-TH");
    }

    function resetFilters() {
      document.getElementById("warehouseFilter").value = "ทั้งหมด";
      document.getElementById("categoryFilter").value = "ทั้งหมด";
      document.getElementById("statusFilter").value = "ทั้งหมด";
      updateDashboard();
    }

    function openAddModal() {
      document.getElementById("modalTitle").textContent = "เพิ่มข้อมูลสินค้า";
      document.getElementById("productForm").reset();
      document.getElementById("editIndex").value = "";
      document.getElementById("reservedQty").value = 0;
      document.getElementById("inTransitQty").value = 0;
document.getElementById("weightKg").value = 0;
      document.getElementById("productModal").classList.add("show");
    }

    function openEditModal(index) {
      const item = stockData[index];

      document.getElementById("modalTitle").textContent = "แก้ไขข้อมูลสินค้า";
      document.getElementById("editIndex").value = index;
      document.getElementById("sku").value = item.sku;
      document.getElementById("productName").value = item.productName;
      document.getElementById("category").value = item.category;
      document.getElementById("warehouse").value = item.warehouse;
      document.getElementById("unitCost").value = item.unitCost;
      document.getElementById("stockOnHand").value = item.stockOnHand;
document.getElementById("weightKg").value = item.weightKg || 0;
      document.getElementById("reservedQty").value = item.reservedQty;
      document.getElementById("inTransitQty").value = item.inTransitQty;
      document.getElementById("reorderPoint").value = item.reorderPoint;
      document.getElementById("maxStock").value = item.maxStock;
      document.getElementById("avgDailySales").value = item.avgDailySales;

      document.getElementById("productModal").classList.add("show");
    }

    function closeModal() {
      document.getElementById("productModal").classList.remove("show");
    }

    function deleteItem(index) {
      const item = stockData[index];
      const confirmDelete = confirm(
        `ต้องการลบสินค้า "${item.productName}" (${item.sku}) ใช่หรือไม่?`
      );

      if (!confirmDelete) return;

      stockData.splice(index, 1);
      saveData();
      updateDashboard();
    }

    document.getElementById("productForm").addEventListener("submit", function(event) {
      event.preventDefault();

      const editIndex = document.getElementById("editIndex").value;
      const sku = document.getElementById("sku").value.trim();

      const newItem = {
        sku,
        productName: document.getElementById("productName").value.trim(),
        category: document.getElementById("category").value.trim(),
        warehouse: document.getElementById("warehouse").value.trim(),
        unitCost: Number(document.getElementById("unitCost").value),
        stockOnHand: Number(document.getElementById("stockOnHand").value),
weightKg: Number(document.getElementById("weightKg").value || 0),
productName: document.getElementById("productName").value.trim(),
        reservedQty: Number(document.getElementById("reservedQty").value || 0),
        inTransitQty: Number(document.getElementById("inTransitQty").value || 0),
        reorderPoint: Number(document.getElementById("reorderPoint").value),
        maxStock: Number(document.getElementById("maxStock").value),
        avgDailySales: Number(document.getElementById("avgDailySales").value)
      };

      const duplicateSku = stockData.some((item, index) =>
        item.sku.toLowerCase() === sku.toLowerCase() &&
        String(index) !== String(editIndex)
      );

      if (duplicateSku) {
        alert("SKU นี้มีอยู่แล้ว กรุณาใช้ SKU อื่น");
        return;
      }

      if (newItem.reservedQty > newItem.stockOnHand) {
        alert("จำนวนที่ถูกจองไม่ควรมากกว่าจำนวนคงเหลือ");
        return;
      }

      if (newItem.maxStock < newItem.reorderPoint) {
        alert("Max Stock ต้องมากกว่าหรือเท่ากับจุดสั่งซื้อ");
        return;
      }

      if (editIndex === "") {
        stockData.push(newItem);
      } else {
        stockData[Number(editIndex)] = newItem;
      }

      saveData();
      closeModal();
      updateDashboard();
    });

    document.getElementById("warehouseFilter").addEventListener("change", updateDashboard);
    document.getElementById("categoryFilter").addEventListener("change", updateDashboard);
    document.getElementById("statusFilter").addEventListener("change", updateDashboard);
function saveTransactions() {
  localStorage.setItem(
    TRANSACTION_STORAGE_KEY,
    JSON.stringify(transactions)
  );
}

function populateTransactionSkuOptions() {
  const select = document.getElementById("transactionSku");

  select.innerHTML = `
    <option value="">-- เลือกสินค้า --</option>
    ${stockData.map((item, index) => `
      <option value="${index}">
        ${item.sku} - ${item.productName} (${item.warehouse})
      </option>
    `).join("")}
  `;
}

function openReceiveModal() {
  document.getElementById("transactionModalTitle").textContent = "รับเข้าสินค้า";
  document.getElementById("transactionSubmitButton").textContent = "บันทึกรับสินค้า";
  document.getElementById("transactionType").value = "รับเข้า";
  document.getElementById("transactionForm").reset();
  document.getElementById("transactionDate").value = new Date().toISOString().split("T")[0];

  populateTransactionSkuOptions();
  document.getElementById("transactionModal").classList.add("show");
}

function openIssueModal() {
  document.getElementById("transactionModalTitle").textContent = "เบิกสินค้า";
  document.getElementById("transactionSubmitButton").textContent = "บันทึกเบิกสินค้า";
  document.getElementById("transactionType").value = "เบิก";
  document.getElementById("transactionForm").reset();
  document.getElementById("transactionDate").value = new Date().toISOString().split("T")[0];

  populateTransactionSkuOptions();
  document.getElementById("transactionModal").classList.add("show");
}

function closeTransactionModal() {
  document.getElementById("transactionModal").classList.remove("show");
}

function renderTransactionTable() {
  const tbody = document.getElementById("transactionTableBody");

  if (transactions.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="8" class="empty">ยังไม่มีประวัติการรับเข้าหรือเบิกสินค้า</td>
      </tr>
    `;
    return;
  }

  const sortedTransactions = [...transactions].sort((a, b) => {
    return new Date(b.createdAt) - new Date(a.createdAt);
  });

  tbody.innerHTML = sortedTransactions.map(item => `
    <tr>
      <td>${new Date(item.date).toLocaleDateString("th-TH")}</td>
      <td class="${item.type === "รับเข้า" ? "transaction-in" : "transaction-out"}">
        ${item.type === "รับเข้า" ? "รับเข้า" : "เบิกสินค้า"}
      </td>
      <td>${item.docNo || "-"}</td>
      <td>${item.sku}</td>
      <td>${item.productName}</td>
      <td>${item.warehouse}</td>
      <td class="number">${numberFormat.format(item.qty)}</td>
      <td>${item.remark || "-"}</td>
    </tr>
  `).join("");
}

document.getElementById("transactionForm").addEventListener("submit", function(event) {
  event.preventDefault();

  const type = document.getElementById("transactionType").value;
  const itemIndex = Number(document.getElementById("transactionSku").value);
  const qty = Number(document.getElementById("transactionQty").value);
  const date = document.getElementById("transactionDate").value;
  const docNo = document.getElementById("transactionDocNo").value.trim();
  const remark = document.getElementById("transactionRemark").value.trim();

  if (document.getElementById("transactionSku").value === "") {
    alert("กรุณาเลือกสินค้า");
    return;
  }

  if (!qty || qty <= 0) {
    alert("กรุณาระบุจำนวนให้มากกว่า 0");
    return;
  }

  const product = stockData[itemIndex];
  const availableQty = product.stockOnHand - product.reservedQty;

  if (type === "เบิก" && qty > availableQty) {
    alert(
      `ไม่สามารถเบิกสินค้าได้ เนื่องจากจำนวนพร้อมใช้มีเพียง ${availableQty} หน่วย`
    );
    return;
  }

  if (type === "รับเข้า") {
    product.stockOnHand += qty;
  } else {
    product.stockOnHand -= qty;
  }

  transactions.push({
    id: Date.now(),
    type,
    date,
    docNo,
    sku: product.sku,
    productName: product.productName,
    warehouse: product.warehouse,
    qty,
    remark,
    createdAt: new Date().toISOString()
  });

  saveData();
  saveTransactions();
  closeTransactionModal();
  updateDashboard();
  renderTransactionTable();

  alert(
    type === "รับเข้า"
      ? "บันทึกรับเข้าสินค้าเรียบร้อย"
      : "บันทึกเบิกสินค้าเรียบร้อย"
  );
});

document.getElementById("transactionModal").addEventListener("click", function(event) {
  if (event.target === this) {
    closeTransactionModal();
  }
});
function exportTransactionsCSV() {
  if (transactions.length === 0) {
    alert("ยังไม่มีประวัติรับเข้าหรือเบิกสินค้าให้ Export");
    return;
  }

  const sortedTransactions = [...transactions].sort((a, b) => {
    return new Date(b.createdAt) - new Date(a.createdAt);
  });

  const headers = [
    "วันที่",
    "ประเภท",
    "เลขที่เอกสาร",
    "SKU",
    "ชื่อสินค้า",
    "คลังสินค้า",
    "จำนวน",
    "หมายเหตุ",
    "วันที่บันทึก"
  ];

  const rows = sortedTransactions.map(item => [
    formatDateForExport(item.date),
    item.type === "รับเข้า" ? "รับเข้า" : "เบิกสินค้า",
    item.docNo || "",
    item.sku || "",
    item.productName || "",
    item.warehouse || "",
    item.qty || 0,
    item.remark || "",
    formatDateTimeForExport(item.createdAt)
  ]);

  // ใส่ BOM (\uFEFF) เพื่อให้ Excel แสดงภาษาไทยถูกต้อง
  const csvContent = "\uFEFF" + [
    headers,
    ...rows
  ].map(row =>
    row.map(value => `"${String(value).replace(/"/g, '""')}"`).join(",")
  ).join("\n");

  const blob = new Blob([csvContent], {
    type: "text/csv;charset=utf-8;"
  });

  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  const today = new Date().toISOString().split("T")[0];

  link.href = url;
  link.download = `ประวัติรับเข้า-เบิกสินค้า_${today}.csv`;

  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  URL.revokeObjectURL(url);
}

function formatDateForExport(dateValue) {
  if (!dateValue) return "";

  const date = new Date(dateValue);

  return date.toLocaleDateString("th-TH", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  });
}

function formatDateTimeForExport(dateValue) {
  if (!dateValue) return "";

  const date = new Date(dateValue);

  return date.toLocaleString("th-TH", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit"
  });
}
function exportTransactionsExcel() {
  if (transactions.length === 0) {
    alert("ยังไม่มีประวัติรับเข้าหรือเบิกสินค้าให้ Export");
    return;
  }

  const sortedTransactions = [...transactions].sort((a, b) => {
    return new Date(b.createdAt) - new Date(a.createdAt);
  });

  const exportData = sortedTransactions.map((item, index) => ({
    "ลำดับ": index + 1,
    "วันที่": formatDateForExport(item.date),
    "ประเภท": item.type === "รับเข้า" ? "รับเข้า" : "เบิกสินค้า",
    "เลขที่เอกสาร": item.docNo || "",
    "SKU": item.sku || "",
    "ชื่อสินค้า": item.productName || "",
    "คลังสินค้า": item.warehouse || "",
    "จำนวน": Number(item.qty) || 0,
    "หมายเหตุ": item.remark || "",
    "วันที่บันทึก": formatDateTimeForExport(item.createdAt)
  }));

  // สร้าง Sheet จากข้อมูล
  const worksheet = XLSX.utils.json_to_sheet(exportData);

  // กำหนดความกว้างคอลัมน์
  worksheet["!cols"] = [
    { wch: 8 },
    { wch: 14 },
    { wch: 14 },
    { wch: 18 },
    { wch: 15 },
    { wch: 35 },
    { wch: 20 },
    { wch: 12 },
    { wch: 40 },
    { wch: 22 }
  ];

  // สร้าง Workbook และเพิ่ม Sheet
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "ประวัติรับเข้า-เบิก"
  );

  // ตั้งชื่อไฟล์ตามวันปัจจุบัน
  const today = new Date().toISOString().split("T")[0];
  const fileName = `ประวัติรับเข้า-เบิกสินค้า_${today}.xlsx`;

  // ดาวน์โหลดไฟล์ Excel
  XLSX.writeFile(workbook, fileName);
}

function importStockFromExcel(event) {
  const file = event.target.files[0];

  if (!file) {
    return;
  }

  if (typeof XLSX === "undefined") {
    alert("ไม่พบระบบอ่านไฟล์ Excel กรุณาตรวจสอบการเพิ่ม SheetJS ในส่วน <head>");
    return;
  }

  const confirmImport = confirm(
    "ระบบจะอัปเดตสินค้าที่มี SKU เดิม และเพิ่มสินค้าที่เป็น SKU ใหม่ ต้องการดำเนินการต่อหรือไม่?"
  );

  if (!confirmImport) {
    event.target.value = "";
    return;
  }

  const reader = new FileReader();

  reader.onload = function(e) {
    try {
      const workbook = XLSX.read(e.target.result, {
        type: "array"
      });

      // อ่านข้อมูลจาก Sheet แรกของไฟล์
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      // แปลงข้อมูล Excel เป็น JSON
      const excelRows = XLSX.utils.sheet_to_json(worksheet, {
        defval: ""
      });

      if (excelRows.length === 0) {
        alert("ไม่พบข้อมูลในไฟล์ Excel");
        return;
      }

      let addedCount = 0;
      let updatedCount = 0;
      let skippedCount = 0;

      excelRows.forEach((row) => {
        // รองรับทั้งหัวคอลัมน์ไทยและอังกฤษ
        const sku = String(
          row["SKU"] ||
          row["sku"] ||
          row["รหัสสินค้า"] ||
          ""
        ).trim();

        if (!sku) {
          skippedCount++;
          return;
        }

        const productName = String(
          row["ชื่อสินค้า"] ||
          row["Product Name"] ||
          row["productName"] ||
          ""
        ).trim();

        // ต้องมี SKU และชื่อสินค้า
        if (!productName) {
          skippedCount++;
          return;
        }

        const importedItem = {
          sku: sku,
          productName: productName,
          category: String(
            row["หมวดหมู่"] ||
            row["Category"] ||
            row["category"] ||
            "ไม่ระบุ"
          ).trim(),

          warehouse: String(
            row["คลังสินค้า"] ||
            row["Warehouse"] ||
            row["warehouse"] ||
            "คลังหลัก"
          ).trim(),

          unitCost: convertToNumber(
            row["ราคาต่อหน่วย"] ||
            row["Unit Cost"] ||
            row["unitCost"]
          ),

          stockOnHand: convertToNumber(
            row["จำนวนคงเหลือ"] ||
            row["Stock On Hand"] ||
            row["stockOnHand"]
          ),

          weightKg: convertToNumber(
            row["น้ำหนัก (กก.)"] ||
            row["น้ำหนัก"] ||
            row["Weight (kg)"] ||
            row["weightKg"]
          ),

          reservedQty: convertToNumber(
            row["จำนวนถูกจอง"] ||
            row["Reserved Qty"] ||
            row["reservedQty"]
          ),

          inTransitQty: convertToNumber(
            row["ระหว่างขนส่ง"] ||
            row["In Transit Qty"] ||
            row["inTransitQty"]
          ),

          reorderPoint: convertToNumber(
            row["จุดสั่งซื้อ"] ||
            row["Reorder Point"] ||
            row["reorderPoint"]
          ),

          maxStock: convertToNumber(
            row["สต๊อกสูงสุด"] ||
            row["Max Stock"] ||
            row["maxStock"]
          ),

          avgDailySales: convertToNumber(
            row["ยอดขายเฉลี่ยต่อวัน"] ||
            row["ยอดขายเฉลี่ย"] ||
            row["Avg Daily Sales"] ||
            row["avgDailySales"]
          )
        };

        // ค้นหาสินค้าเดิมโดยอ้างอิง SKU
        const existingIndex = stockData.findIndex((item) => {
          return String(item.sku).trim().toLowerCase() === sku.toLowerCase();
        });

        if (existingIndex !== -1) {
          // SKU เดิม: อัปเดตข้อมูลสินค้า
          stockData[existingIndex] = {
            ...stockData[existingIndex],
            ...importedItem
          };

          updatedCount++;
        } else {
          // SKU ใหม่: เพิ่มสินค้าใหม่
          stockData.push(importedItem);
          addedCount++;
        }
      });

      saveData();
      updateDashboard();

      alert(
        "Import ข้อมูลสำเร็จ\n\n" +
        "เพิ่มสินค้าใหม่: " + addedCount + " รายการ\n" +
        "อัปเดตสินค้าเดิม: " + updatedCount + " รายการ\n" +
        "ข้ามรายการไม่สมบูรณ์: " + skippedCount + " รายการ"
      );

    } catch (error) {
      console.error(error);
      alert("ไม่สามารถอ่านไฟล์ Excel ได้ กรุณาตรวจสอบรูปแบบไฟล์และหัวคอลัมน์");
    } finally {
      // ล้างไฟล์ที่เลือก เพื่อให้สามารถเลือกไฟล์เดิมซ้ำได้
      event.target.value = "";
    }
  };

  reader.readAsArrayBuffer(file);
}

function convertToNumber(value) {
  if (value === null || value === undefined || value === "") {
    return 0;
  }

  // รองรับตัวเลขที่มี comma เช่น 1,250.50
  const numberValue = Number(
    String(value).replace(/,/g, "").trim()
  );

  return Number.isFinite(numberValue) ? numberValue : 0;
}
function exportStockExcel() {
  // Export ตามข้อมูลที่ผ่านตัวกรองปัจจุบัน
  const filteredData = getFilteredData();

  if (filteredData.length === 0) {
    alert("ไม่พบรายการสินค้าให้ Export");
    return;
  }

  // เรียง SKU: SKU-001, SKU-002, SKU-010
  const sortedData = [...filteredData].sort((a, b) => {
    return String(a.sku).localeCompare(String(b.sku), undefined, {
      numeric: true,
      sensitivity: "base"
    });
  });

  // จัดรูปแบบข้อมูลที่จะส่งออก Excel
  const exportData = sortedData.map((item, index) => {
    const stockOnHand = Number(item.stockOnHand) || 0;
    const reservedQty = Number(item.reservedQty) || 0;

    // จำนวนพร้อมใช้ = คงเหลือ - จำนวนถูกจอง
    const availableQty = stockOnHand - reservedQty;

    const unitCost = Number(item.unitCost) || 0;
    const stockValue = stockOnHand * unitCost;

    // ใช้สถานะเดิมของระบบ หากมี หากไม่มีให้คำนวณจากจุดสั่งซื้อ
    let status = item.status || "";

    if (!status) {
      const reorderPoint = Number(item.reorderPoint) || 0;

      if (availableQty <= 0) {
        status = "หมดสต๊อก";
      } else if (availableQty <= reorderPoint) {
        status = "ควรสั่งซื้อ";
      } else {
        status = "ปกติ";
      }
    }

    // จำนวนแนะนำสั่ง
    const maxStock = Number(item.maxStock) || 0;
    const suggestedOrderQty = Math.max(0, maxStock - stockOnHand);

    return {
      "ลำดับ": index + 1,
      "SKU": item.sku || "",
      "ชื่อสินค้า": item.productName || "",
      "หน่วยนับ": item.unit || "ชิ้น",
      "หมวดหมู่": item.category || "",
      "คลังสินค้า": item.warehouse || "",
      "จำนวนคงเหลือ": stockOnHand,
      "จำนวนถูกจอง": reservedQty,
      "จำนวนพร้อมใช้": availableQty,
      "น้ำหนัก (กก.)": Number(item.weightKg) || 0,
      "จุดสั่งซื้อ": Number(item.reorderPoint) || 0,
      "แนะนำสั่ง": suggestedOrderQty,
      "ราคาต่อหน่วย": unitCost,
      "มูลค่าสต๊อก": stockValue,
      "สถานะ": status
    };
  });

  // แปลงข้อมูลเป็น Excel Worksheet
  const worksheet = XLSX.utils.json_to_sheet(exportData);

  // กำหนดความกว้างคอลัมน์
  worksheet["!cols"] = [
    { wch: 8 },   // ลำดับ
    { wch: 15 },  // SKU
    { wch: 35 },  // ชื่อสินค้า
    { wch: 12 },  // หน่วยนับ
    { wch: 18 },  // หมวดหมู่
    { wch: 18 },  // คลังสินค้า
    { wch: 15 },  // จำนวนคงเหลือ
    { wch: 15 },  // จำนวนถูกจอง
    { wch: 15 },  // จำนวนพร้อมใช้
    { wch: 15 },  // น้ำหนัก
    { wch: 14 },  // จุดสั่งซื้อ
    { wch: 14 },  // แนะนำสั่ง
    { wch: 16 },  // ราคาต่อหน่วย
    { wch: 18 },  // มูลค่าสต๊อก
    { wch: 16 }   // สถานะ
  ];

  // สร้าง Workbook
  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "รายการสินค้า"
  );

  // ตั้งชื่อไฟล์ด้วยวันที่ปัจจุบัน
  const today = new Date().toISOString().split("T")[0];

  XLSX.writeFile(
    workbook,
    `รายการสินค้าคงเหลือ_${today}.xlsx`
  );
}
    document.getElementById("productModal").addEventListener("click", function(event) {
      if (event.target === this) closeModal();
    });

    updateDashboard();
  </script>
</body>
</html>
