// index.js
const fs = require("fs");
const Excel = require("exceljs");

let raworder = fs.readFileSync("package.json");
let jsonorder = JSON.parse(raworder);

jsonorder.sort((a, b) => new Date(a.OrderDate) - new Date(b.OrderDate));

let initialPurchasesRaworder = fs.readFileSync("sample.json");
let initialPurchases = JSON.parse(initialPurchasesRaworder);

initialPurchases = initialPurchases.reduce((acc, purchase) => {
  const key = `${purchase.skuId}_${purchase.Purchase}`;
  acc[key] = { ...purchase, balance: parseInt(purchase["Total Inbound"], 10) };
  return acc;
}, {});

let workbook = new Excel.Workbook();
let worksheet = workbook.addWorksheet("My Sheet");

worksheet.addRow([
  "OrderDate",
  "OrderId",
  "ProductId",
  "SkuId",
  "ProductLink",
  "3RDLink",
  "ProductImageUrl",
  "Quantity",
  "Price",
  "SubTotal",
  "ConversionRate",
  "VariationName_1",
  "VariationValue_1",
  "VariationImage_1",
  "VariationName_2",
  "VariationValue_2",
  "VariationImage_2",
  "Commission",
  "SellerId",
  "Height",
  "Weight",
  "Width",
  "Length",
  "Name",
  "PurchasingDate",
  "ItemStatus",
  "MappingPurchaseID",
  "TotalQuantityRev",
  "NoteOrCancelReason",
  "PurchasingPrice",
  "CustomerName",
  "PhoneNumber",
  "Salesman",
  "OrderStatus",
  "PriceFromSaboPerItem",
  "PurchasingSubtotal",
  "TotalPurchaseDeliveryFee",
  "TotalPurchaseServiceFee",
  "Discount",
  "FinalPurchaseTotalAmount",
  "City",
  "ProductRename",
  "CheckPurchaseID",
  "CancelReason",
  "ActualDelivered",
  "ReturnQuantity",
  "CanCalculationPurchasing",
  "NeedCheckPurchaseId",
  "DeliveryFee",
  "FinalQuantity",
  "SourceComponentItem",
  "ReturnQuantity",
  "ReturnReason",
  "ReturnDate",
  "FinalSubtotal",
  "DeliveryFeeByItem",
  "StockAvailable",
]);

function writeToWorksheet(order, purchaseId, quantity, remainingQuantity) {}

jsonorder.forEach((order) => {
  worksheet.addRow([
    order.OrderDate,
    order.OrderId,
    order.ProductId,
    order.SkuId,
    order.ProductLink,
    order["3 RD link"],
    order.ProductImageUrl,
    order.Quantity,
    order.Price,
    order.SubTotal,
    order["Conversion Rate"],
    order.VariationName_1,
    order.VariationValue_1,
    order.VariationImage_1,
    order.VariationName_2,
    order.VariationValue_2,
    order.VariationImage_2,
    order.Commission,
    order.SellerId,
    order.Height,
    order.Weight,
    order.Width,
    order.Length,
    order.Name,
    order["Purchasing Date"],
    order.ItemStatus,
    order["Mapping Purchase ID"],
    order["Total Quantity Rev"],
    order["Note or Cancel Reason"],
    order["Purchasing Price"],
    order["Customer Name"],
    order["Phone number"],
    order.Salesman,
    order["Order Status"],
    order["Price From Sabo Per Item"],
    order["Purchasing subtotal"],
    order["Total Purchase delivery fee"],
    order["Total purchase service fee"],
    order.Discount,
    order["Final Purchase Total Amount"],
    order.City,
    order.ProductRename,
    order["Check Purchase ID"],
    order["Cancel Reason"],
    order["Actual Delivered"],
    order["Return Quatity"],
    order["CanCaculationPurchasing"],
    order["NeedCheckPurchaseId"],
    order.DeliveryFee,
    order["FinalQuantity"],
    order.SourceComponentItem,
    order["ReturnQuantity"],
    order["ReturnReason"],
    order["ReturnDate"],
    order["FinalSubtotal"],
    order.DeliverfyFeeByItem,
    order.StockAvailable,
    "", // Vì không có giá trị cho trường cuối cùng
  ]);
  let sku = order.SkuId;
  let quantityNeeded = parseInt(order.Quantity, 10);
  let orderHandled = false;

  // Sắp xếp possiblePurchases theo thứ tự DateTime để ưu tiên lấy hàng từ purchase cũ hơn
  let possiblePurchases = Object.keys(initialPurchases)
    .filter((key) => key.startsWith(`${sku}_`))
    .sort(
      (a, b) =>
        new Date(initialPurchases[a].DateTime) -
        new Date(initialPurchases[b].DateTime)
    );

  for (let purchaseKey of possiblePurchases) {
    let purchaseorder = initialPurchases[purchaseKey];
    let availableQuantity = purchaseorder.balance;

    if (availableQuantity >= quantityNeeded) {
      // Đủ số lượng trong purchase này
      purchaseorder.balance -= quantityNeeded;
      writeToWorksheet(
        order,
        purchaseorder.Purchase,
        quantityNeeded,
        purchaseorder.balance
      );
      orderHandled = true;
      break;
    } else if (availableQuantity > 0) {
      // Lấy toàn bộ số lượng có sẵn từ purchase này và tiếp tục với purchase khác
      quantityNeeded -= availableQuantity;
      purchaseorder.balance = 0;
      writeToWorksheet(order, purchaseorder.Purchase, availableQuantity, 0);
    }
  }

  if (!orderHandled && quantityNeeded > 0) {
    // Trường hợp không đủ số lượng trong tất cả các purchase
    // Xử lý theo yêu cầu cụ thể, có thể báo lỗi hoặc ghi nhận việc thiếu hàng
    console.log(`Không đủ hàng cho SkuId: ${sku}, thiếu ${quantityNeeded}`);
  }
});

// Ghi file
workbook.xlsx.writeFile("test-excel.xlsx").then(function () {
  console.log("File saved!");
});
