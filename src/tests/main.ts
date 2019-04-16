import { Workbook } from './../index';

const workbook = new Workbook('./','text.xlsx');
const sheet = workbook.createSheet('test',2,2);
sheet.set(1,1,'Yes');
workbook.save().then(()=>{
    // tslint:disable-next-line:no-console
    console.log('done');
});