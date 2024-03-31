import csv from "csvtojson";
import fs from "fs";
import {
  chain,
  cloneDeep,
  groupBy,
  mapKeys,
  mapValues,
  min,
  sortBy,
  sumBy,
} from "lodash";
import ExcelJS from "exceljs";

// read json file to get the data for todays change
const todaysChange = JSON.parse(
  fs.readFileSync("./todaysChange.json", "utf8") || "{}"
);

const profitStep = 25000;

type Order = {
  fund: string;
  date: Date;
  transaction_type: string;
  quantity: number;
  amount: number;
  last_price: number;
  average_price: number;
  now: number;
  profit: number;
};

const splitOrder = (
  order: Order,
  a: number,
  b: number
): [Order, Order] | [null, Order] | [Order, null] => {
  if (Math.abs(a) < Math.abs(b * Number.EPSILON)) return [null, order];
  if (Math.abs(b) < Math.abs(a * Number.EPSILON)) return [order, null];
  if (a < 0) return [null, order];
  if (b < 0) return [order, null];

  const ratio = a / (a + b);
  const orderA = { ...order };
  const orderB = { ...order };

  orderA.quantity = Math.round(orderA.quantity * ratio * 1000) / 1000;
  orderA.amount = Math.round(orderA.average_price * orderA.quantity);
  orderA.now = Math.round(orderA.last_price * orderA.quantity);
  orderA.profit = orderA.now - orderA.amount;

  orderB.quantity = order.quantity - orderA.quantity;
  orderB.amount = order.amount - orderA.amount;
  orderB.now = order.now - orderA.now;
  orderB.profit = orderB.now - orderB.amount;

  return [orderA, orderB];
};

const addMilestones = (
  orders: Order[],
  key: keyof { [P in keyof Order as Order[P] extends number ? P : never]: P },
  milestone: number,
  labelOn?: string,
  labelNext?: string
): [boolean, number] => {
  let index = -1;
  let sum = 0;

  do {
    index++;
    sum += orders[index][key];
  } while (
    milestone - sum > milestone * Number.EPSILON &&
    index < orders.length - 1
  );

  if (sum >= milestone) {
    const milestoneOrder = orders[index];
    const extra = sum - milestone;
    const split = splitOrder(
      milestoneOrder,
      milestoneOrder[key] - extra,
      extra
    );

    orders.splice(index, 1, ...[...split].filter((x): x is Order => !!x));

    if (!split[0]) index--;

    if (orders[index] && labelOn) orders[index].fund += ` (${labelOn})`;
    if (orders[index + 1] && labelNext)
      orders[index + 1].fund += ` (${labelNext})`;
    return [true, index];
  }

  return [false, -1];
};

const fillSheetWithOrders = (
  worksheet: ExcelJS.Worksheet,
  ordersByFund: Record<string, Order[]>,
  funds: string[],
  extraCalculations: boolean
) => {
  worksheet.getRow(1).values = [
    "Fund",
    "Date",
    "Quantity",
    "Last Price",
    "Average Price",
    "Amount",
    "Now",
    "Profit/Loss",
  ];
  worksheet.getRow(1).font = { bold: true, size: 12 };

  let row = 4;

  funds.forEach((fund) => {
    let fundOrders = ordersByFund[fund];
    fundOrders.forEach((order) => {
      const values = [
        order.fund,
        order.date,
        order.quantity,
        order.last_price,
        order.average_price,
        order.amount,
        order.now,
        order.profit,
      ];
      const Row = worksheet.getRow(row);
      Row.font = { size: 12 };
      Row.values = values;
      // Row.getCell("C").numFmt = "#0.000";
      Row.getCell("F").numFmt = "#,##0;(#,##0)";
      Row.getCell("G").numFmt = "#,##0;(#,##0)";
      Row.getCell("H").numFmt = "#,##0;(#,##0)";
      row++;
    });

    if (extraCalculations) {
      const xirrRow = worksheet.getRow(row - fundOrders.length);
      const xirrCell = xirrRow.getCell("I");
      xirrCell.value = {
        formula: `XIRR(F${row - fundOrders.length}:F${row},B${
          row - fundOrders.length
        }:B${row}, 0.2)`,
        date1904: false,
      };
      xirrCell.font = { bold: true, size: 12 };
      xirrCell.numFmt = "0.00%";

      const profitRow = worksheet.getRow(row - fundOrders.length + 1);
      const profitCell = profitRow.getCell("I");
      profitCell.value = {
        formula: `SUM(H${row - fundOrders.length}:H${row - 1})`,
        date1904: false,
      };
      profitCell.font = { bold: true, size: 12 };
      profitCell.numFmt = "#,##0;(#,##0)";
    }
    const totalRow = worksheet.getRow(row);
    totalRow.font = { size: 12 };
    totalRow.getCell(1).value = "Total";
    totalRow.getCell(1).font = { bold: true, size: 12 };
    totalRow.getCell(2).value = new Date();
    totalRow.getCell(6).value = {
      formula: `-ROUND(SUM(C${row - fundOrders.length}:C${row - 1})*D${
        row - 1
      },0)`,
      date1904: false,
    };
    totalRow.getCell(6).font = { bold: true, size: 12 };
    totalRow.getCell(6).numFmt = "#,##0;(#,##0)";
    row += 4;
  });

  if (extraCalculations) {
    worksheet.getRow(5).getCell("M").value = "XIRR";
    worksheet.getRow(5).getCell("M").font = { bold: true, size: 12 };
    worksheet.getRow(6).getCell("M").value = "Profit/Loss";
    worksheet.getRow(6).getCell("M").font = { bold: true, size: 12 };
    worksheet.getRow(7).getCell("M").value = "Total invested";
    worksheet.getRow(7).getCell("M").font = { bold: true, size: 12 };

    worksheet.getRow(5).getCell("N").value = {
      formula: `XIRR(F4:F${row - 4},B4:B${row - 4}, 0.2)`,
      date1904: false,
    };
    worksheet.getRow(5).getCell("N").font = { bold: true, size: 12 };
    worksheet.getRow(5).getCell("N").numFmt = "0.00%";
    worksheet.getRow(6).getCell("N").value = {
      formula: `-SUM(F4:F${row - 4})`,
      date1904: false,
    };
    worksheet.getRow(6).getCell("N").font = { bold: true, size: 12 };
    worksheet.getRow(6).getCell("N").numFmt = "#,##0;(#,##0)";
    worksheet.getRow(7).getCell("N").value = {
      formula: `SUM(G4:G${row - 4})+SUM(F4:F${row - 4})`,
      date1904: false,
    };
    worksheet.getRow(7).getCell("N").font = { bold: true, size: 12 };
    worksheet.getRow(7).getCell("N").numFmt = "#,##0;(#,##0)";
  }
  // set column width to accomodate widest cell in each column
  worksheet.columns.forEach((column) => {
    let maxLength = 0;
    column.eachCell?.({ includeEmpty: false }, (cell) => {
      // get cell value as string, if it's a date then format it as YYYY-MM-DD
      const value =
        cell.value instanceof Date ? "YYYY-MM-DD" : cell.value || "";
      const columnLength = value.toString().length;
      if (columnLength > maxLength) {
        maxLength = columnLength;
      }
    });
    column.width = Math.max(maxLength, 10);
  });
};

const build = async () => {
  const data: {
    order_id: string;
    exchange_order_id: string;
    tradingsymbol: string;
    status: string;
    status_message: string;
    folio: string;
    fund: string;
    order_timestamp: string;
    exchange_timestamp: string;
    settlement_id: string;
    transaction_type: string;
    variety: string;
    purchase_type: string;
    quantity: string;
    amount: string;
    last_price: string;
    average_price: string;
    placed_by: string;
    payment_confirmed: string;
    fund_source: string;
    tag: string;
  }[] = await csv().fromFile("./data.csv");

  const orders: Order[] = chain(data)
    .filter((x) => x.status.toLowerCase() === "complete")
    .sortBy((x) => x.fund)
    .sortBy((x) => x.order_timestamp)
    .map((item) => {
      todaysChange[item.fund] = todaysChange[item.fund] || 0;
      const lastPriceFactor = (100 + todaysChange[item.fund]) / 100;

      const factor = item.transaction_type === "BUY" ? 1 : -1;

      const quantity = parseFloat(item.quantity) * factor;
      // const amount = Math.round(parseFloat(item.amount)) * factor;
      const amount = item.transaction_type === "BUY" ? Math.round(parseFloat(item.average_price) * quantity *100) / 100 : Math.round(parseFloat(item.amount)) * factor ;
      const last_price =
        Math.round(parseFloat(item.last_price) * lastPriceFactor * 10000) /
        10000;
      const now = Math.round(last_price * quantity);

      return {
        fund: item.fund,
        date: new Date(item.order_timestamp.split("T")[0]),
        transaction_type: item.transaction_type,
        quantity,
        amount,
        last_price,
        average_price: parseFloat(item.average_price),
        now,
        profit: now - amount,
      };
    })
    .value();

  fs.writeFileSync(
    "./todaysChange.json",
    JSON.stringify(todaysChange, null, 2),
    "utf8"
  );

  const ordersByFund = groupBy(cloneDeep(orders), (x) => x.fund);
  const funds = chain(ordersByFund)
    .keys()
    .sortBy(
      (x) =>
        chain(ordersByFund[x])
          .minBy((y) => y.date)
          .value().date
    )
    .value();

  funds.forEach((fund) => {
    let fundOrders = ordersByFund[fund];

    fundOrders = chain(fundOrders)
      .sortBy((x) => x.date)
      .reverse()
      .sortBy((x) => x.transaction_type)
      .reverse()
      .value();

    const [held, holdingIndex] = addMilestones(
      fundOrders,
      "quantity",
      0,
      "Sold till here",
      "Holding start"
    );

    if (held) {
      const heldOrders = fundOrders.slice(holdingIndex + 1);
      const totalProfit = chain(heldOrders)
        .sumBy((x) => x.profit)
        .value();
      for (
        let target = profitStep;
        target <= totalProfit;
        target += profitStep
      ) {
        addMilestones(heldOrders, "profit", target, `Profit reached ${target}`);
      }
      fundOrders.splice(holdingIndex + 1, heldOrders.length, ...heldOrders);
    }

    ordersByFund[fund] = fundOrders;
  });

  const workbook = new ExcelJS.Workbook();

  fillSheetWithOrders(
    workbook.addWorksheet("By Date", {
      views: [{ state: "frozen", ySplit: 1 }],
    }),
    groupBy(orders, (x) => x.fund),
    funds,
    true
  );

  fillSheetWithOrders(
    workbook.addWorksheet("By Holdings", {
      views: [{ state: "frozen", ySplit: 1 }],
    }),
    ordersByFund,
    funds,
    false
  );

  await workbook.xlsx.writeFile("orders.xlsx");
};

build();
