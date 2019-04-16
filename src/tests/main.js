const { Workbook } = require("./../../lib/index");

const workbook = new Workbook("./", "text.xlsx");
const sheet = workbook.createSheet("test", 5, 10);
workbook.save().then(() => {
  console.log("done");
});
