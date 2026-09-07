function transferOrder(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orders = ss.getSheetByName("Orders");
  const values = e.values;

  // Find the first empty row using the Order Date column
  const maxRows = orders.getMaxRows();
  const orderDates = orders.getRange(2, 2, maxRows - 1, 1).getValues();
  const emptyIndex = orderDates.findIndex(row => row[0] === "");

  let nextRow;

  if (emptyIndex === -1) {
    orders.insertRowAfter(maxRows);
    nextRow = maxRows + 1;
  } else {
    nextRow = emptyIndex + 2;
  }

  // Get the date as a real date value from Form Responses
  const responseSheet = e.range.getSheet();
  const responseRow = e.range.getRow();
  const orderDate = responseSheet.getRange(responseRow, 2).getValue();

  // B:G — Order Date to Unit Price
  orders.getRange(nextRow, 2, 1, 6).setValues([[
    orderDate,
    values[2],
    String(values[3]),
    values[4],
    Number(values[5]),
    Number(values[6])
  ]]);

  // I — Amount Received
  orders.getRange(nextRow, 9).setValue(Number(values[7]) || 0);

  // K — Payment Method
  orders.getRange(nextRow, 11).setValue(values[8]);

  // M:O — Order Status, Sales Channel, Notes
  orders.getRange(nextRow, 13, 1, 3).setValues([[
    values[9],
    values[10],
    values[11] || ""
  ]]);
}
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("Mobile Order Entry");
}

function getProducts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lists = ss.getSheetByName("Lists");
  const lastRow = lists.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  return lists
    .getRange(2, 1, lastRow - 1, 1)
    .getDisplayValues()
    .flat()
    .filter(product => product !== "");
}

function submitWebsiteOrder(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orders = ss.getSheetByName("Orders");
    const savedPin = PropertiesService
  .getScriptProperties()
  .getProperty("BUSINESS_PIN");
  
const enteredPin = String(data.pin || "").trim();

if (!savedPin || enteredPin !== savedPin) {
  throw new Error("Incorrect security PIN.");
}

    const customerName = String(data.customerName || "").trim();
    const mobile = String(data.mobile || "").replace(/\D/g, "");
    const product = String(data.product || "").trim();
    const quantity = Number(data.quantity);
    const unitPrice = Number(data.unitPrice);
    const amountReceived = Number(data.amountReceived);

    if (!customerName) {
      throw new Error("Customer name is required.");
    }

    if (!/^[6-9][0-9]{9}$/.test(mobile)) {
      throw new Error("Enter a valid 10-digit mobile number.");
    }

    if (!product) {
      throw new Error("Select a product.");
    }

    if (!(quantity > 0)) {
      throw new Error("Quantity must be greater than zero.");
    }

    if (!(unitPrice > 0)) {
      throw new Error("Unit price must be greater than zero.");
    }

    if (amountReceived < 0) {
      throw new Error("Amount received cannot be negative.");
    }

    const dateParts = String(data.orderDate).split("-").map(Number);

    if (dateParts.length !== 3) {
      throw new Error("Enter a valid order date.");
    }

    const orderDate = new Date(
      dateParts[0],
      dateParts[1] - 1,
      dateParts[2]
    );

    const maximumRows = orders.getMaxRows();
    const existingDates = orders
      .getRange(2, 2, maximumRows - 1, 1)
      .getValues();

    const emptyIndex = existingDates.findIndex(row => row[0] === "");
    let nextRow;

    if (emptyIndex === -1) {
      orders.insertRowAfter(maximumRows);
      nextRow = maximumRows + 1;
    } else {
      nextRow = emptyIndex + 2;
    }

    orders.getRange(nextRow, 2, 1, 6).setValues([[
      orderDate,
      customerName,
      mobile,
      product,
      quantity,
      unitPrice
    ]]);

    orders.getRange(nextRow, 9).setValue(amountReceived);
    orders.getRange(nextRow, 11).setValue(data.paymentMethod);

    orders.getRange(nextRow, 13, 1, 3).setValues([[
      data.orderStatus,
      data.salesChannel,
      String(data.notes || "").trim()
    ]]);

    SpreadsheetApp.flush();

    const totalAmount = quantity * unitPrice;
    const pendingAmount = totalAmount - amountReceived;
    const orderId = orders.getRange(nextRow, 1).getDisplayValue();

    return {
      success: true,
      orderId: orderId,
      totalAmount: totalAmount,
      pendingAmount: pendingAmount
    };
  } finally {
    lock.releaseLock();
  }
}
