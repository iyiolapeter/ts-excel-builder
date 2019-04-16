import { Workbook } from "./../index";

const workbook = new Workbook("./", "text1.xlsx");
const sheet = workbook.createSheet("test", 5, 10);
sheet.set(1,1,'Entry Reference');
sheet.set(2,1,'Beneficiary Number');
sheet.set(3,1,'Amount');
sheet.set(4,1,'Payable');
sheet.set(5,1,'Fees');
workbook.save().then(() => {
  // tslint:disable-next-line:no-console
  console.log("done");
});
