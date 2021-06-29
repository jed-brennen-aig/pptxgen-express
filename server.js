const pptxgen = require("pptxgenjs");
const got = require("got");
const express = require("express");
require("dotenv").config();
const app = express();

const redisClient = require("./redis-client");
const mockData = require("./mock-data");

const spacingBorder = [
  { type: "none" },
  { type: "none" },
  { type: "solid", color: "ffffff", pt: 5 },
  { type: "none" },
];

const convertTableBodyRows = (tableData) => {
  const pptRows = [];
  const colCount = tableData.thead[0].cells.length;

  tableData.tbody &&
    tableData.tbody.forEach((row) => {
      if ("title" in row) {
        pptRows.push([
          {
            text: row.title,
            options: {
              colspan: colCount,
              fill: { color: "e8ecf1" },
              border: spacingBorder,
            },
          },
        ]);
        row.rows.forEach((groupRow) => {
          pptRows.push(groupRow.cells.map((cell) => ({ text: cell.value })));
        });
      } else if ("cells" in row) {
        return pptRows.push(row.cells.map((cell) => ({ text: cell.value })));
      }
    });
  return pptRows;
};

const convertTableHeaderRows = (tableData) => {
  const pptRows = [];

  tableData.thead &&
    tableData.thead.forEach((header) =>
      pptRows.push(
        header.cells.map((cell) => ({
          text: cell.value,
          options: {
            border: spacingBorder,
            bold: true,
          },
        }))
      )
    );

  return pptRows;
};

const generatePowerpoint = (data) => {
  const pptx = new pptxgen();

  for (let i = 0; i < 10; i++) {
    const rows = [
      ...convertTableHeaderRows(data),
      ...convertTableBodyRows(data),
    ];

    pptx
      .addSlide({ masterName: "DS.CML_PRODUCTION_INVESTMENT_ACTIVITY" })
      .addTable(rows, {
        align: "left",
        fontFace: "Source Sans Pro",
        autoPage: true,
      });

    let dataChartAreaLine = [
      {
        name: "Actual Sales",
        labels: [
          "Jan",
          "Feb",
          "Mar",
          "Apr",
          "May",
          "Jun",
          "Jul",
          "Aug",
          "Sep",
          "Oct",
          "Nov",
          "Dec",
        ],
        values: [
          1500, 4600, 5156, 3167, 8510, 8009, 6006, 7855, 12102, 12789, 10123,
          15121,
        ],
      },
      {
        name: "Projected Sales",
        labels: [
          "Jan",
          "Feb",
          "Mar",
          "Apr",
          "May",
          "Jun",
          "Jul",
          "Aug",
          "Sep",
          "Oct",
          "Nov",
          "Dec",
        ],
        values: [
          1000, 2600, 3456, 4567, 5010, 6009, 7006, 8855, 9102, 10789, 11123,
          12121,
        ],
      },
    ];

    pptx.addSlide().addChart(pptx.ChartType.line, dataChartAreaLine, {
      x: 1,
      y: 1,
      w: 8,
      h: 4,
    });
  }

  return pptx.stream();
};

const getPptxData = async () => {
  const cacheKey = "pptx_data";
  const cacheValue = await redisClient.getAsync(cacheKey);
  if (!cacheValue) {
    console.log("FETCHING");
    const promise = new Promise((resolve, reject) => {
      setTimeout(() => resolve(mockData), 3000);
    });
    const pptxData = await promise;
    redisClient.setAsync(cacheKey, JSON.stringify(pptxData));

    return pptxData;
  }
  console.log("USING CACHE");

  return JSON.parse(cacheValue);
};

const getPowerPoint = async (req, res, next) => {
  try {
    const data = await getPptxData();
    const pptx = await generatePowerpoint(data.mockTableData);

    res.writeHead(200, {
      "Content-disposition": "attachment;filename=Test.pptx",
      "Content-Length": pptx.length,
    });
    res.end(Buffer.from(pptx, "binary"));
  } catch (error) {
    next(error);
  }
};

app.get("/powerpoint", getPowerPoint);

app.get("/", (req, res) => {
  return res.send("Hello world");
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
});
