import fs from 'fs';
import JSZip from 'jszip';
import path from 'path';
import xml from 'xmlbuilder';

interface StringObject {
    [key:string]:string
}

const baseXl:StringObject = {
    '_rels/.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    'docProps/core.xml':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Administrator</dc:creator><cp:lastModifiedBy></cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2006-09-13T11:21:51Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2006-09-13T11:21:55Z</dcterms:modified></cp:coreProperties>',
    'xl/theme/theme1.xml':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office 主题"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:FillStyleDefinitionLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:FillStyleDefinitionLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleDefinitionLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleDefinitionLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>',
    'xl/styles.xml':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><color theme="1"/><name val="宋体"/><family val="2"/><charset val="134"/><scheme val="minor"/></font><font><sz val="9"/><name val="宋体"/><family val="2"/><charset val="134"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf></cellXfs><cellStyles count="1"><cellStyle name="常规" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/></styleSheet>'
}

const tool = {
    i2a: (i:number)=>{
        return 'ABCDEFGHIJKLMNOPQRSTUVWXYZ123'.charAt(i-1)
    }
}

const opt = {
    tmpl_path : __dirname
}

abstract class Xmler {

    public book: Workbook;
    constructor(book: Workbook){
        this.book = book;
    }
    public abstract toxml():string;
}

class ContentTypes extends Xmler {

    public toxml(){
        const types = xml.create('Types',{version:'1.0',encoding:'UTF-8',standalone:true})
        types.att('xmlns','http://schemas.openxmlformats.org/package/2006/content-types');
        types.ele('Override',{PartName:'/xl/theme/theme1.xml',ContentType:'application/vnd.openxmlformats-officedocument.theme+xml'})
        types.ele('Override',{PartName:'/xl/styles.xml',ContentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'})
        types.ele('Default',{Extension:'rels',ContentType:'application/vnd.openxmlformats-package.relationships+xml'})
        types.ele('Default',{Extension:'xml',ContentType:'application/xml'})
        types.ele('Override',{PartName:'/xl/workbook.xml',ContentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'})
        types.ele('Override',{PartName:'/docProps/app.xml',ContentType:'application/vnd.openxmlformats-officedocument.extended-properties+xml'})
        const length = this.book.sheets.length;
        for (let i = 1; i <= length; i++){
            types.ele('Override',{PartName:'/xl/worksheets/sheet'+i+'.xml',ContentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'})
        }
        types.ele('Override',{PartName:'/xl/sharedStrings.xml',ContentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'})
        types.ele('Override',{PartName:'/docProps/core.xml',ContentType:'application/vnd.openxmlformats-package.core-properties+xml'})
        const Xml = types.end();
        return Xml;
    }
}



class SharedStrings {
    public cache:{[key:string]:number} = {};
    public arr:string[] = [];

    public str2id(s:string){
        if(this.cache[s]){
            return this.cache[s];
        }
        this.cache[s] = this.arr.push(s);
        return this.cache[s];
    }

    public toxml(){
        const sst = xml.create('sst',{version:'1.0',encoding:'UTF-8',standalone:true})
        sst.att('xmlns','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        sst.att('count',''+this.arr.length)
        sst.att('uniqueCount',''+this.arr.length)
        for (const element of this.arr){
            const si = sst.ele('si')
            si.ele('t',element)
            si.ele('phoneticPr',{fontId:1,type:'noConversion'})
        }
        const Xml = sst.end();
        return Xml;
    }
}

class DocPropsApp extends Xmler {
    
    public toxml(){
        const props = xml.create('Properties',{version:'1.0',encoding:'UTF-8',standalone:true})
        props.att('xmlns','http://schemas.openxmlformats.org/officeDocument/2006/extended-properties')
        props.att('xmlns:vt','http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes')
        props.ele('Application','Microsoft Excel')
        props.ele('DocSecurity','0')
        props.ele('ScaleCrop','false')
        let tmp = props.ele('HeadingPairs').ele('vt:vector',{size:2,baseType:'variant'})
        tmp.ele('vt:variant').ele('vt:lpstr','工作表')
        tmp.ele('vt:variant').ele('vt:i4',''+this.book.sheets.length)
        tmp = props.ele('TitlesOfParts').ele('vt:vector',{size:this.book.sheets.length,baseType:'lpstr'});
        const length = this.book.sheets.length;
        for (let i = 1; i <= length; i++){
            tmp.ele('vt:lpstr',this.book.sheets[i-1].name);
        }
        props.ele('Company')
        props.ele('LinksUpToDate','false')
        props.ele('SharedDoc','false')  
        props.ele('HyperlinksChanged','false')  
        props.ele('AppVersion','12.0000') 
        return props.end()
    }
}

class XlWorkbook extends Xmler {
    
    public toxml(){
        const wb = xml.create('workbook',{version:'1.0',encoding:'UTF-8',standalone:true})
        wb.att('xmlns','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        wb.att('xmlns:r','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        wb.ele('fileVersion',{appName:'xl',lastEdited:'4',lowestEdited:'4',rupBuild:'4505'})
        wb.ele('workbookPr',{filterPrivacy:'1',defaultThemeVersion:'124226'}) 
        wb.ele('bookViews').ele('workbookView',{xWindow:'0',yWindow:'90',windowWidth:'19200',windowHeight:'11640'})
        const tmp = wb.ele('sheets');
        const length = this.book.sheets.length;
        for (let i = 1; i <= length; i++){
            tmp.ele('sheet',{name:this.book.sheets[i-1].name,sheetId:''+i,'r:id':'rId'+i});
        }
        wb.ele('calcPr',{calcId:'124519'})
        return wb.end();
    }
}

class XlRels extends Xmler {
    
    public toxml(){
        const rs = xml.create('Relationships',{version:'1.0',encoding:'UTF-8',standalone:true})
        rs.att('xmlns','http://schemas.openxmlformats.org/package/2006/relationships')
        const length = this.book.sheets.length;
        for (let i = 1; i <= length; i++){
            rs.ele('Relationship',{Id:'rId'+i,Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',Target:'worksheets/sheet'+i+'.xml'})
        }
        rs.ele('Relationship',{Id:'rId'+(length+1),Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',Target:'theme/theme1.xml'})
        rs.ele('Relationship',{Id:'rId'+(length+2),Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',Target:'styles.xml'})
        rs.ele('Relationship',{Id:'rId'+(length+3),Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',Target:'sharedStrings.xml'})
        return rs.end()
    }
}

export class Workbook {
    public id:string = '';
    public sheets:Sheet[] = [];
    public ss:SharedStrings;
    public ct:ContentTypes;
    public da:DocPropsApp;
    public wb:XlWorkbook;
    public re:XlRels;
    public st:Style;

    public fpath: string;
    public fname: string;

    constructor(fpath:string, fname:string){
        this.id = String(Math.floor(Math.random()*9999999));
        this.ct = new ContentTypes(this);
        this.ss = new SharedStrings();
        this.da = new DocPropsApp(this);
        this.wb = new XlWorkbook(this);
        this.re = new XlRels(this);
        this.st = new Style(this);
        this.fpath = fpath;
        this.fname = fname;
    }

    public createSheet(name:string, cols:number, rows:number){
        const sheet = new Sheet(this, name, cols, rows);
        this.sheets.push(sheet);
        return sheet;
    }

    public save(){
        return new Promise((resolve,reject)=>{
            const target = path.resolve(this.fpath,this.fname);
            this.generate().then((zip)=>{
                return zip.generateAsync({type: 'nodebuffer'});
            }).then((buffer)=>{
                fs.writeFile(target, buffer, (error)=>{
                    if(error){
                        reject(error);
                    }
                    resolve(true);
                });
            }).catch((error)=>{
                reject(error);
            })
        });
    }

    private generate(){
        return new Promise<JSZip>((resolve,reject)=>{
            try {
                const zip = new JSZip;
                // tslint:disable-next-line:forin
                for(const key in baseXl){
                    zip.file(key,baseXl[key]);
                }
                // # 1 - build [Content_Types].xml
                zip.file('[Content_Types].xml',this.ct.toxml())
                // # 2 - build docProps/app.xml
                zip.file('docProps/app.xml',this.da.toxml())
                // # 3 - build xl/workbook.xml
                zip.file('xl/workbook.xml',this.wb.toxml())
                // # 4 - build xl/sharedStrings.xml
                zip.file('xl/sharedStrings.xml',this.ss.toxml())
                // # 5 - build xl/_rels/workbook.xml.rels
                zip.file('xl/_rels/workbook.xml.rels',this.re.toxml())
                // # 6 - build xl/worksheets/sheet(1-N).xml
                const length = this.sheets.length;
                for (let i = 1; i <= length; i++){
                    zip.file('xl/worksheets/sheet'+(i)+'.xml',this.sheets[i-1].toxml())
                }
                // # 7 - build xl/styles.xml
                zip.file('xl/styles.xml',this.st.toxml())
                resolve(zip);
            } catch (error) {
                reject(error);
            }
        });
    }

}

export interface CellDefinition {
    col: number,
    row: number
}

export interface BorderStyleDefinition {
    left: string,
    right: string,
    top: string,
    bottom: string
}

export interface FontStyleDefinition {
    name?: string,
    sz?: string,
    color?: string,
    family?: string,
    scheme?: string,
    bold?: string,
    iter?: string
}

export interface FillStyleDefinition {
    type: string,
    fgColor: string,
    bgColor: string
}

export interface StyleDefinition {
    align: string,
    valign: string,
    rotate: string,
    wrap?: string,
    font_id: number,
    fill_id: number,
    bder_id: number
}

export class Sheet {
    public name:string;
    public rows: number;
    public cols: number;
    private book:Workbook;
    private data:{[key:number]:{[key:number]:{v:number}}} = {};
    private merges:Array<{from: CellDefinition, to: CellDefinition}> = [];
    private col_wd:Array<{c:number,cw:number}> = [];
    private row_ht:{[key:number]:number} = {};
    private styles:{[key:string]:number | string} = {};
    constructor(book:Workbook, name:string, cols:number, rows:number){
        this.book = book;
        this.name = name;
        this.data = {};
        for(let i=0; i < rows; i++){
            this.data[i+1] = {}
            for(let j=0; j < cols; j++){
                this.data[i+1][j+1] = {v:0};
            }
        }
        this.rows = rows;
        this.cols = cols;
    }

    public set(col: number, row: number, data: string){
        this.data[row][col].v = this.book.ss.str2id(data);
    }

    public merge(from:CellDefinition, to:CellDefinition){
        this.merges.push({from,to});
    }

    public width(col:number, wd:number){
        this.col_wd.push({c:col, cw: wd});
    }

    public height(row: number, height: number){
        this.row_ht[row] = height;
    }

    public font(col: number, row:number, font_s: FontStyleDefinition){
        this.styles['font_'+col+'_'+row] = this.book.st.font2id(font_s);
    }

    public fill(col: number, row: number, fill_s: FillStyleDefinition){
        this.styles['fill_'+col+'_'+row] = this.book.st.fill2id(fill_s);
    }

    public border(col: number, row: number, bder_s: BorderStyleDefinition){
        this.styles['bder_'+col+'_'+row] = this.book.st.bder2id(bder_s)
    }

    public align(col: number, row: number, align_s: string){
        this.styles['algn_'+col+'_'+row] = align_s;
    }

    public valign(col:number, row: number, valign_s: string){
        this.styles['valgn_'+col+'_'+row] = valign_s
    }

    public rotate(col: number, row: number, angle: string){
        this.styles['rotate_'+col+'_'+row] = angle;
    }

    public wrap(col: number, row: number, wrap_s: string){
        this.styles['wrap_'+col+'_'+row] = wrap_s
    }

    public style_id(col: number, row: number){
        const inx = '_'+col+'_'+row;
        const definition: StyleDefinition = {
            font_id: Number(this.styles['font'+inx]),
            fill_id: Number(this.styles['fill'+inx]),
            bder_id: Number(this.styles['bder'+inx]),
            align: String(this.styles['algn'+inx]),
            valign: String(this.styles['valgn'+inx]),
            rotate: String(this.styles['rotate'+inx]),
            wrap: String(this.styles['wrap'+inx]) 
        };
        return this.book.st.style2id(definition);
    }

    public toxml(){
        const ws = xml.create('worksheet',{version:'1.0',encoding:'UTF-8',standalone:true});
        ws.att('xmlns','http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        ws.att('xmlns:r','http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        ws.ele('dimension',{ref:'A1'});
        ws.ele('sheetViews').ele('sheetView',{workbookViewId:'0'});
        ws.ele('sheetFormatPr',{defaultRowHeight:'13.5'});
        if (this.col_wd.length > 0){
            const cols = ws.ele('cols')
            for (const cw of this.col_wd){
                cols.ele('col',{min:''+cw.c,max:''+cw.c,width:cw.cw,customWidth:'1'})
            }
        }
        const sd = ws.ele('sheetData');
        for (let i=1; i <= this.rows; i++){
            const r = sd.ele('row',{r:''+i,spans:'1:'+this.cols})
            const ht = this.row_ht[i]
            if(ht){
                r.att('ht',ht)
                r.att('customHeight','1')
            }        
            for (let j=1; j <= this.cols; j++){
                const ix = this.data[i][j];
                const sid = this.style_id(j,i);
                if (ix.v > 0 || sid !== 1){
                    const c = r.ele('c',{r:''+tool.i2a(j)+i});
                    if(sid !== 1) { c.att('s',''+(sid-1)); }
                    if (ix.v !== 0) {
                        c.att('t','s')
                    }
                        c.ele('v',''+(ix.v-1))
                    }
            }
        }
        if (this.merges.length > 0) {
            const mc = ws.ele('mergeCells',{count:this.merges.length});
            for(const m of this.merges){
                mc.ele('mergeCell',{ref:(''+tool.i2a(m.from.col)+m.from.row+':'+tool.i2a(m.to.col)+m.to.row)});
            }
        }
        ws.ele('phoneticPr',{fontId:'1',type:'noConversion'});
        ws.ele('pageMargins',{left:'0.7',right:'0.7',top:'0.75',bottom:'0.75',header:'0.3',footer:'0.3'});
        ws.ele('pageSetup',{paperSize:'9',orientation:'portrait',horizontalDpi:'200',verticalDpi:'200'});
        return ws.end();
    }
}

class Style extends Xmler{

    public mfonts:FontStyleDefinition[] = [];
    public mfills:FillStyleDefinition[] = [];
    public mbders:BorderStyleDefinition[] = [];
    public mstyle:StyleDefinition[] = [];

    public def_font_id:number;
    public def_fill_id:number;
    public def_bder_id:number;
    public def_style_id:number;

    public def_align:string = '-';
    public def_valign:string = '-';
    public def_rotate:string = '-';
    public def_wrap:string = '-';
    
    private cache:{[key:string]:number}={};

    private def_font: FontStyleDefinition = {
        name: '宋体',
        sz: '11',
        color: '-',
        family: '2',
        scheme: 'minor',
        bold: '-',
        iter: '-'
    };
    private def_fill: FillStyleDefinition = {
        type: 'none',
        bgColor: '-',
        fgColor: '-'
    };
    private def_bder: BorderStyleDefinition = {
        left: '-',
        right: '-',
        top: '-',
        bottom: '-'
    };

    constructor(book: Workbook){
        super(book);
        this.def_font_id = this.font2id();
        this.def_fill_id = this.fill2id();
        this.def_bder_id = this.bder2id();
        this.def_style_id = this.style2id({
            font_id:this.def_font_id,
            fill_id:this.def_fill_id,
            bder_id:this.def_bder_id,
            align:this.def_align,
            valign:this.def_valign,
            rotate:this.def_rotate,
            wrap:this.def_wrap
        });
    }

    public font2id(font?:FontStyleDefinition){
        font = Object.assign(this.def_font,font?font:{});
        const k = 'font_'+font.bold+font.iter+font.sz+font.color+font.name+font.scheme+font.family;
        if(!this.cache[k]){
            this.cache[k] = this.mfonts.push(font);
        }
        return this.cache[k];
    }

    public fill2id(fill?:FillStyleDefinition){
        fill = Object.assign(this.def_fill,fill?fill:{});
        const k = 'fill_' + fill.type + fill.bgColor + fill.fgColor;
        if(!this.cache[k]){
            this.cache[k] = this.mfills.push(fill);
        }
        return this.cache[k];
    }

    public bder2id(bder?:BorderStyleDefinition){
        bder = Object.assign(this.def_bder,bder?bder:{});
        const k = 'bder_'+bder.left+'_'+bder.right+'_'+bder.top+'_'+bder.bottom;
        if(!this.cache[k]){
            this.cache[k] = this.mbders.push(bder);
        }
        return this.cache[k];
    }

    public style2id(style:StyleDefinition){
        const k = 's_' + style.font_id + '_' + style.fill_id + '_' + style.bder_id + '_' + style.align + '_' + style.valign + '_' + style.wrap + '_' + style.rotate;
        if(!this.cache[k]){
            this.cache[k] = this.mstyle.push(style);
        }
        return this.cache[k];
    }

    public toxml() {
        const ss = xml.create('styleSheet',{version:'1.0',encoding:'UTF-8',standalone:true})
        ss.att('xmlns','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        const fonts = ss.ele('fonts',{count:this.mfonts.length})
        for (const o of this.mfonts){
            const e = fonts.ele('font');
            if(o.bold !== '-') { e.ele('b'); }
            if(o.iter !== '-') { e.ele('i'); }
            e.ele('sz',{val:o.sz})
            if(o.color !== '-') { e.ele('color',{theme:o.color}); }
            e.ele('name',{val:o.name})
            e.ele('family',{val:o.family})
            e.ele('charset',{val:'134'})
            if(o.scheme !== '-') { e.ele('scheme',{val:'minor'}); }
        }
        const fills = ss.ele('fills',{count:this.mfills.length})
        for (const o of this.mfills){
            const e = fills.ele('fill')
            const es = e.ele('patternFill',{patternType:o.type})
            if(o.fgColor !== '-') { es.ele('fgColor',{theme:'8',tint:'0.79998168889431442'}); }
            if(o.bgColor !== '-') { es.ele('bgColor',{indexed:o.bgColor}); }
        }
        const bders = ss.ele('borders',{count:this.mbders.length})
        for (const o of this.mbders){
            const e = bders.ele('border');
            (o.left !== '-')?e.ele('left',{style:o.left}).ele('color',{auto:'1'}):e.ele('left');
            (o.right !== '-')?e.ele('right',{style:o.right}).ele('color',{auto:'1'}):e.ele('right');
            (o.top !== '-')?e.ele('top',{style:o.top}).ele('color',{auto:'1'}):e.ele('top');
            (o.bottom !== '-')?e.ele('bottom',{style:o.bottom}).ele('color',{auto:'1'}):e.ele('bottom');
            e.ele('diagonal');
        }
        ss.ele('cellStyleXfs',{count:'1'}).ele('xf',{numFmtId:'0',fontId:'0',fillId:'0',borderId:'0'}).ele('alignment',{vertical:'center'})
        const cs = ss.ele('cellXfs',{count:this.mstyle.length});
        for (const o of this.mstyle) {
            const e = cs.ele('xf',{numFmtId:'0',fontId:(o.font_id-1),fillId:(o.fill_id-1),borderId:(o.bder_id-1),xfId:'0'})
            if(o.font_id !== 1) { e.att('applyFont','1'); }
            if(o.fill_id !== 1) { e.att('applyFill','1'); }
            if(o.bder_id !== 1) { e.att('applyBorder','1'); }
            if (o.align !== '-' || o.valign !== '-' || o.wrap !== '-'){
                e.att('applyAlignment','1')
                const ea = e.ele('alignment',{textRotation:((o.rotate === '-')?'0':o.rotate),horizontal:((o.align === '-')?'left':o.align), vertical:((o.valign === '-')?'top':o.valign)});
                if (o.wrap !== '-') { ea.att('wrapText','1'); }
            }
        }
        ss.ele('cellStyles',{count:'1'}).ele('cellStyle',{name:'常规',xfId:'0',builtinId:'0'})
        ss.ele('dxfs',{count:'0'})
        ss.ele('tableStyles',{count:'0',defaultTableStyle:'TableStyleMedium9',defaultPivotStyle:'PivotStyleLight16'})
        return ss.end();
    }

}