const { Workbook } = require("./../lib/index");

const workbook = new Workbook("./", "text.xlsx");
const sheet = workbook.createSheet("test", 2, 2);
workbook.save().then(() => {
  console.log("done");
});
