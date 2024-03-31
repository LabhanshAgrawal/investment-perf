import csv from "csvtojson";
import { chain, cloneDeep, groupBy, maxBy } from "lodash";
import ExcelJS from "exceljs";

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

const fillSheetWithOrders = (
  worksheet: ExcelJS.Worksheet,
  ordersByFund: Record<string, Order[]>,
  funds: string[]
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
    trade_date: string;
    trade_type: string;
    quantity: string;
    price: string;
    strike: string;
    order_id: string;
    trade_id: string;
    series: string;
    exchange: string;
    segment: string;
    order_execution_time: string;
    isin: string;
    instrument_id: string;
    instrument_type: string;
    tradingsymbol: string;
    expiry_date: string;
    external_trade_type: string;
    tag_ids: string;
  }[] = await csv().fromFile("./data2.csv");

  // create a map with keys as trading symbol and values as the latest price from data
  const latestPrices = chain(data)
    .groupBy("tradingsymbol")
    .mapValues((x) => {
      const latest = maxBy(x, (y) => y.trade_date)!;
      return parseFloat(latest.price);
    })
    .value();

  const orders: Order[] = chain(data)
    .sortBy((x) => x.tradingsymbol)
    .sortBy((x) => x.trade_date)
    .map((item) => {
      const factor = item.trade_type === "buy" ? 1 : -1;

      const quantity = parseFloat(item.quantity) * factor;
      const amount = parseFloat(item.price) * quantity;
      const last_price = latestPrices[item.tradingsymbol];
      const now = last_price * quantity;

      return {
        fund: item.tradingsymbol,
        date: new Date(item.trade_date),
        transaction_type: item.trade_type,
        quantity,
        amount,
        last_price,
        average_price: parseFloat(item.price),
        now,
        profit: now - amount,
      };
    })
    .value();

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

  const workbook = new ExcelJS.Workbook();

  fillSheetWithOrders(
    workbook.addWorksheet("By Date", {
      views: [{ state: "frozen", ySplit: 1 }],
    }),
    groupBy(orders, (x) => x.fund),
    funds
  );

  await workbook.xlsx.writeFile("kite.xlsx");
};

build();
