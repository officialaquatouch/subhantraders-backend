const express = require("express");
const cors = require("cors");
const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

const app = express();
const PORT = 5000;
const EXCEL_FILE = path.join(__dirname, "orders", "MS_Iron_Orders.xlsx");

app.use(cors());
app.use(express.json());

if (!fs.existsSync(path.join(__dirname, "orders"))) {
  fs.mkdirSync(path.join(__dirname, "orders"));
}

async function initExcel() {
  if (!fs.existsSync(EXCEL_FILE)) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Orders");
    const headers = ["Order #","Date","Time","Name","Phone","City","Product","Qty","Unit Price","Subtotal","COD","Total","Note","Status"];
    const headerRow = ws.addRow(headers);
    const widths   = [10,18,12,22,18,16,30,8,14,14,12,14,30,14];
    widths.forEach((w,i) => { ws.getColumn(i+1).width = w; });
    headerRow.eachCell(cell => {
      cell.fill = { type:"pattern", pattern:"solid", fgColor:{ argb:"FF10107A" } };
      cell.font = { bold:true, color:{ argb:"FFFFCC00" }, size:11, name:"Calibri" };
      cell.alignment = { vertical:"middle", horizontal:"center" };
    });
    headerRow.height = 28;
    await wb.xlsx.writeFile(EXCEL_FILE);
    console.log("✅ Excel file created!");
  }
}

async function saveOrderToExcel(order) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(EXCEL_FILE);
  const ws = wb.getWorksheet("Orders");
  const orderNo = ws.rowCount;
  const now = new Date();
  const dateStr = now.toLocaleDateString("en-PK", { day:"2-digit", month:"short", year:"numeric" });
  const timeStr = now.toLocaleTimeString("en-PK", { hour:"2-digit", minute:"2-digit" });
  const unitPrice = parseInt(order.unit_price) || 0;
  const qty = parseInt(order.quantity) || 1;
  const subtotal = unitPrice * qty;
  const cod = subtotal > 0 ? 350 : 0;
  const total = subtotal + cod;

  const newRow = ws.addRow([
    orderNo,
    dateStr,
    timeStr,
    order.name,
    order.phone,
    order.city || "—",
    order.product,
    qty,
    unitPrice > 0 ? `Rs. ${unitPrice.toLocaleString()}` : "Wholesale",
    subtotal > 0 ? `Rs. ${subtotal.toLocaleString()}` : "—",
    cod > 0 ? `Rs. ${cod}` : "—",
    total > 0 ? `Rs. ${total.toLocaleString()}` : "Contact",
    order.note || "—",
    "Pending"
  ]);

  const isEven = orderNo % 2 === 0;
  newRow.eachCell(cell => {
    cell.fill = { type:"pattern", pattern:"solid", fgColor:{ argb: isEven ? "FFE8E8F5" : "FFFFFFFF" } };
    cell.alignment = { vertical:"middle", horizontal:"center", wrapText:true };
    cell.font = { name:"Calibri", size:10 };
  });

  // Total cell bold
  newRow.getCell(12).font = { bold:true, color:{ argb:"FF10107A" }, size:11, name:"Calibri" };

  // Status cell
  newRow.getCell(14).fill = { type:"pattern", pattern:"solid", fgColor:{ argb:"FFFFF3CD" } };
  newRow.getCell(14).font = { bold:true, color:{ argb:"FF856404" }, size:10, name:"Calibri" };

  newRow.height = 22;
  await wb.xlsx.writeFile(EXCEL_FILE);
  return { orderNo, total };
}

app.post("/api/order", async (req, res) => {
  try {
    const { name, phone, city, product, unit_price, quantity, note } = req.body;
    if (!name || !phone) {
      return res.status(400).json({ success:false, message:"Name and phone required." });
    }
    const result = await saveOrderToExcel({ name, phone, city, product, unit_price, quantity, note });
    console.log(`✅ Order #${result.orderNo} — ${name} — ${phone}`);
    res.json({ success:true, message:"Order saved!", order_no:result.orderNo, total:result.total });
  } catch (err) {
    console.error("❌ Error:", err);
    res.status(500).json({ success:false, message:"Server error." });
  }
});

app.get("/api/orders", async (req, res) => {
  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(EXCEL_FILE);
    const ws = wb.getWorksheet("Orders");
    const orders = [];
    ws.eachRow((row, rowNum) => {
      if (rowNum === 1) return;
      orders.push({
        order_no: row.getCell(1).value,
        date:     row.getCell(2).value,
        name:     row.getCell(4).value,
        phone:    row.getCell(5).value,
        city:     row.getCell(6).value,
        product:  row.getCell(7).value,
        total:    row.getCell(12).value,
        status:   row.getCell(14).value,
      });
    });
    res.json({ success:true, count:orders.length, orders });
  } catch (err) {
    res.status(500).json({ success:false, message:"Could not read orders." });
  }
});

app.get("/api/download", (req, res) => {
  if (!fs.existsSync(EXCEL_FILE)) {
    return res.status(404).json({ message:"No orders yet." });
  }
  res.download(EXCEL_FILE, "MS_Iron_Orders.xlsx");
});

app.get("/", (req, res) => {
  res.json({ status:"MS Iron Backend Running ✅" });
});

initExcel().then(() => {
  app.listen(PORT, () => {
    console.log(`\n🚀 Backend running at http://localhost:${PORT}\n`);
  });
});