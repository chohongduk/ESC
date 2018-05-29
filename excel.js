const mysql = require('mysql');
var express = require('express');
var excelforjson = require('excel4node');
var excel = require('exceljs');
var bodyParser = require('body-parser');
var app = express();
var multer = require('multer');
var storage = multer.diskStorage({
  destination: function(req, file, cb){
    cb(null, 'C:/Users/hufscse/Desktop/gaja')
  },
  filename: function(req, file, cb){
    cb(null, file.originalname)
  }
})
var upload = multer({storage: storage})

if(typeof require !== 'undefined') XLSX = require('xlsx');


var conn = mysql.createConnection({
  host     : 'localhost',
  user     : 'root',
  password : 'qufwhdemf',
  database : 'excel'
});



conn.connect();

app.set('view engine', 'jade');
app.use(bodyParser.urlencoded({extended:false}))
app.use('/', express.static('public'));

app.get('/', function(req,res){
  res.render('file');
})
app.get('/change', function (req,res) {
  res.render('change')
})
app.post('/insert', function(req,res){
  console.log('sucess');
  res.render('change');
})
app.post('/view_data',function(req,res){
  var SelectCompany = req.body.selected;
  console.log(SelectCompany);
  var qw=[];
  var responseData=[];
  var count1;
  var loadData1 = function(){
    return new Promise((resolve, reject)=>{
    conn.query(`select TABLE_ROWS from information_schema.tables where table_name =  '${SelectCompany}'`,function(error,resu,fields){

      count1 = resu[0].TABLE_ROWS;
      console.log(count1);
      if(count1 != undefined) resolve('Success');
    });



    })

    }

  loadData1().then((val)=>{
    cnt=count1;
    for(var i = 0; i<count1; i++){
        (function(closed_i){
        conn.query(`select distinct NAME from ${SelectCompany}`,function(error, results, fields){
          if(error) reject('failed');
        //  console.log(results[closed_i])
          qw[closed_i]=results[closed_i];
          cnt--;
          if(qw[closed_i] !== undefined){
        //    console.log(qw[closed_i].NAME);
           responseData[closed_i]=qw[closed_i];
          }
          if(cnt===0){
            // console.log(responseData);
            res.json(responseData);

          }
        });
      }) (i);

    }


  }).catch((err)=>{
    console.log(err);
  });
})
app.post('/view_gyu', function(req,res){
  var SelectCompany = req.body.selected1;
  var Selectname = req.body.selected;
  var ab = "count(gyu)";
  var qw=[];
  var responseData=[];
  var count1;
  var loadData2 = function(){
    return new Promise((resolve, reject)=>{
    conn.query(`select gyu from ${SelectCompany} where name = '${Selectname}'`,function(error,resu,fields){
      count1 = resu.length;
      console.log(resu);
      if(count1 != undefined) {
        resolve('Success');
        res.json(resu);
      }
    });



    })

    }
    loadData2().then((val)=>{
      console.log(count1);


    }).catch((err)=>{
      console.log(err);
    });
});
app.post('/view_price',function(req,res){
  var SelectCompany = req.body.selected1;
  var Selectname = req.body.selected;
  var Selectgyu = req.body.selected2;
  var Selectprice = req.body.selected3;
  var a = "";
  Selectprice = parseInt(Selectprice);

  console.log(Selectprice);

  var loadData3 = function(){
    return new Promise((resolve, reject)=>{
    conn.query(`select material,labor,public from ${SelectCompany} where name = '${Selectname}' and gyu = '${Selectgyu}'`,function(error,resu,fields){

      resu[0].material=parseInt(resu[0].material);
      resu[0].labor=parseInt(resu[0].labor);
      resu[0].public=parseInt(resu[0].public);
      if(resu[0].material > 0){
        console.log(resu[0].material);
        a = "material";
      }
      else if(resu[0].labor > 0)
        a = "labor";
      else if(resu[0].public > 0)
        a = "public";

      console.log(a);
      if(resu != undefined) {
        resolve('Success');
      }
    });



    })

    }
    loadData3().then((val)=>{

      conn.query(`update ${SelectCompany} set ${a} = ${Selectprice} where name = '${Selectname}' and gyu = '${Selectgyu}'`,function(error,resu,fields){

      })


    }).catch((err)=>{
      console.log(err);
    });
})
app.post('/insert_data', function(req,res){

  var a = req.body.cho;
  var exjson = XLSX.readFile(a);
  var sheet_name_list = exjson.SheetNames;
  var json = XLSX.utils.sheet_to_json(exjson.Sheets[sheet_name_list[1]]);
  var length = json.length;
  // console.log(json[1]);
  for(i=0;i<length;i++){
  var q1 = json[i].품목코드;
  var q2 = json[i].품명;
  var q3 = json[i].규격;
  var q4 = json[i].재료비;
  var q5 = json[i].노무비;
  var q6 = json[i].경비;
  var q7 = json[i].공산품;
  var q8 = json[i].광산품;
  var q9 = json[i].공사노임;
  var q10 = json[i].전력수도,도시가스;
  var q11 = json[i].기타;
  var q12 = json[i].외산;
  var q13 = json[i].국산;

  q7=parseInt(q7);
  q8=parseInt(q8);
  q9=parseInt(q9);
  q10=parseInt(q10);
  q11=parseInt(q11);
  q12=parseInt(q12);
  q13=parseInt(q13);

    if(q7>0) {
      var t = '공산품';
    }
    else if(q8>0) {
      var t = '광산품';
    }
    else if(q9>0) {
      var t = '공사노임';
    }
    else if(q10>0) {
      var t = '전력수도,도시가스';
    }
    else if(q11>0) {
      var t = '기타';
    }
    else if(q12>0) {
      var t = '외산';
    }
    else if(q13>0) {
      var t = '국산';
    }



  conn.query(`insert into excel.test2 values('${q1}', '${q2}','${q3}', '${q4}', '${q5}', '${q6}', '${t}')`, function (error, results, fields){

  })
}
  res.render('')
})
app.post('/form_receiver', upload.single('cho'), function(req,res){
  var a = req.file.originalname;
  // var b = req.body.title;
  // console.log(req.file);
  // console.log(b);
  var as = req.body.db;
  // console.log(as);


// var exjson = XLSX.readFile('haja.xlsx');
var exjson = XLSX.readFile(a);
var sheet_name_list = exjson.SheetNames;
var json = XLSX.utils.sheet_to_json(exjson.Sheets[sheet_name_list[0]]);
var length = json.length;



var writeworkbook = new excel.Workbook();//새로운 엑셀파일 지정
var wb = new excelforjson.Workbook();
var ws = wb.addWorksheet('내역서');
var sheet = writeworkbook.addWorksheet('내역서');
var sheet1 = writeworkbook.addWorksheet('빈내역서');
var worksheet = writeworkbook.getWorksheet('내역서');
var worksheet1 = writeworkbook.getWorksheet('빈내역서');
function create(){

  worksheet.columns = [
    {header :'', key : 'number'},
    {header :'소계', key : 'so'},
    {header :'품목코드', key : 'code', width : 10},
    {header :'품명', key : 'name', width : 10},
    {header :'규격', key : 'size', width : 10},
    {header :'단위', key : 'unit', width : 10},
    {header :'재료비', key : 'material', width : 10},
    {header :'노무비', key : 'labor', width : 10},
    {header :'경비', key : 'public', width : 10},
    {header :'수량', key : 'count', width : 10},
    {header :'재료비', key: 'materialwon'},
    {header :'노무비', key: 'laborwon'},
    {header :'경비', key: 'publicwon'},
    {header :'공사노임', key : 'noim', width : 10},
    {header :'국산', key : 'domestic', width : 10},
    {header :'외산', key : 'foreign', width : 10},
    {header :'기타', key : 'etc', width : 10},
    {header :'광산품', key : 'industry', width : 10},
    {header :'공산품', key : 'mine', width : 10},
    {header :'전력수도,도시가스', key : 'gas', width : 10},
    {header :'농림수산품', key : 'nong', width : 10},
    {header :'토목부문', key: 'G1'},
    {header :'건축부문', key: 'G2'},
    {header :'기계부문', key: 'G3'},
    {header :'전기부문', key: 'G4'},
    {header :'합계', key: 'total'},
    {header :'비고', key: 'bigo'},
  ]

  worksheet1.columns = [
    {header :'', key : 'number'},
    {header :'소계', key : 'so'},
    {header :'품목코드', key : 'code', width : 10},
    {header :'품명', key : 'name', width : 10},
    {header :'규격', key : 'size', width : 10},
    {header :'단위', key : 'unit', width : 10},
    {header :'재료비', key : 'material', width : 10},
    {header :'노무비', key : 'labor', width : 10},
    {header :'경비', key : 'public', width : 10},
    {header :'수량', key : 'count', width : 10},
    {header :'재료비', key: 'materialwon'},
    {header :'노무비', key: 'laborwon'},
    {header :'경비', key: 'publicwon'},
    {header :'공사노임', key : 'noim', width : 10},
    {header :'국산', key : 'domestic', width : 10},
    {header :'외산', key : 'foreign', width : 10},
    {header :'기타', key : 'etc', width : 10},
    {header :'광산품', key : 'industry', width : 10},
    {header :'공산품', key : 'mine', width : 10},
    {header :'전력수도,도시가스', key : 'gas', width : 10},
    {header :'농림수산품', key : 'nong', width : 10},
    {header :'토목부문', key: 'G1'},
    {header :'건축부문', key: 'G2'},
    {header :'기계부문', key: 'G3'},
    {header :'전기부문', key: 'G4'},
    {header :'합계', key: 'total'},
    {header :'비고', key: 'bigo'},
  ]

};



var trg = '*';
var tblName = 'test2';
var a = [];

var loadData = function(){
  return new Promise((resolve, reject)=>{
    var cnt = length;
    for(var i = 0; i<length; i++){

      (function(closed_i){
        var mame = json[closed_i].품목코드;
        conn.query(`select * from ${as} where code = '${mame}'`, function (error, results, fields){
          if (error) reject('failed');
          a[closed_i] = results[0];
          if(results.length < 1) console.log(closed_i, results);
          // console.log(a[closed_i]);
          cnt--;
          if( cnt === 0 ) resolve('Success');
        });
      })(i);
    }
  });
}

create();
var font1 = { name : 'HY견명조', size: 10};
var font2 = { name : 'Arial Narrow', size: 10};
var fill1 = {
  type: 'pattern',
  pattern: 'solid',
  fgColor:{argb:'FF808080'},
  bgColor:{argb:'FF0000FF'}
}
var fill2 = {
  type: 'pattern',
  pattern: 'solid',
  fgColor:{argb:'FFFFE400'},
  bgColor:{argb:'FF0000FF'}
}
var border1 = {
  top: {style:'dotted'},
  left: {style:'dotted'},
  bottom: {style:'dotted'},
  right: {style:'dotted'}
}
var border2 = {
  top: {style:'dotted'},
  left: {style:'thin'},
  bottom: {style:'dotted'},
  right: {style:'thin'}
}
var border3 = {
  top: {style:'dotted'},
  left: {style:'thin'},
  bottom: {style:'dotted'},
  right: {style:'thick', color : {argb : 'FF0000FF'}}
}
var border4 = {
  top: {style:'dotted'},
  left: {style:'thin'},
  bottom: {style:'dotted'},
  right: {style:'dotted'}
}
worksheet.getCell('A1').value = '';

worksheet.getCell('B1').value = '';
worksheet.getCell('C1').value = '';
worksheet.getCell('D1').value = '';
worksheet.getCell('E1').value = '';
worksheet.getCell('F1').value = '';
worksheet.getCell('G1').value = '';
worksheet.getCell('H1').value = '';
worksheet.getCell('I1').value = '';
worksheet.getCell('J1').value = '';
worksheet.getCell('K1').value = '';
worksheet.getCell('L1').value = '';
worksheet.getCell('M1').value = '';
worksheet.getCell('N1').value = '';
worksheet.getCell('O1').value = '';
worksheet.getCell('P1').value = '';
worksheet.getCell('Q1').value = '';
worksheet.getCell('R1').value = '';
worksheet.getCell('S1').value = '';
worksheet.getCell('T1').value = '';
worksheet.getCell('U1').value = '';
worksheet.getCell('V1').value = '';
worksheet.getCell('W1').value = '';
worksheet.getCell('X1').value = '';
worksheet.getCell('Y1').value = '';
worksheet.getCell('Z1').value = '';
worksheet.getCell('AA1').value = '';
worksheet.getCell('AB1').value = '';
worksheet.getCell('D2').value = '◈ 공사명 : HUFS 조 and 정 Capstone';
worksheet.getCell('D2').font=font1;

worksheet.addRow(['','소계','품목','','','','단  가','','','금  액','','','','비 목 군 분 류 금 액','','','','','','','','','','','','']).alignment = { vertical: 'middle', horizontal: 'center'};
worksheet.addRow(['','CODE','CODE','품  명','규격','단위','재료비','노무비','경비','수량','재료비','노무비','경비','공사노임','국산','외산','기타','광산품','공산품(D)','전력수도,도시가스','농림수산품','','','','','합계','비고']).alignment = { vertical: 'middle', horizontal: 'center' };
worksheet.addRow(['','','','','','','','','','','','','','(A)','(B1)','(B2)','(Z)','(C)','(D)','(E)','(F)','(토목부문G1)','(건축부문G2)','(기계부문G3)','(전기부문G4)']).alignment = { vertical: 'middle', horizontal: 'center' };
worksheet.addRow([1,'<<자재>>','','<<자재>>']);
worksheet.mergeCells('D2:H2');
worksheet.mergeCells('G3:I3');
worksheet.mergeCells('J3:M3');
worksheet.mergeCells('N3:Z3');
worksheet.getRow(3).font = font1;
worksheet.getRow(4).font = font1;
worksheet.getRow(5).font = font1;
worksheet.getRow(1).fill = fill1;
worksheet.getRow(6).fill = fill2;
worksheet.getRow(2).border={
  top:{style:'thick', color : {argb : 'FF0000FF'}},
  bottom:{style:'thin'}
}






loadData().then((val)=> {
  for(var i=0; i<length;i++){
    if(a[i] === undefined){
      worksheet.addRow({number: i+2,code: json[i].품목코드, name: json[i].품명, size: json[i].규격, unit: json[i].단위, count: json[i].수량}).font=font2;
      worksheet1.addRow({number: i+2,code: json[i].품목코드, name: json[i].품명, size: json[i].규격,unit: json[i].단위, count: json[i].수량}).font=font2;
      continue;
    }
    if(a[i].bimoc === '공산품'){
      var oi = a[i].material * json[i].수량;
      worksheet.addRow({number: i+2,code: json[i].품목코드, name: json[i].품명,size: json[i].규격,unit: json[i].단위,count: json[i].수량,materialwon: oi,material: a[i].material, labor: a[i].labor, public: a[i].public, mine: a[i].material}).font=font2;
    }
    else if (a[i].bimoc === '기타') {
      var oi = a[i].material * json[i].수량;
      worksheet.addRow({number: i+2,code: json[i].품목코드, name: json[i].품명,size: json[i].규격,unit: json[i].단위,count: json[i].수량,materialwon: oi, material: a[i].material, labor: a[i].labor, public: a[i].public, etc: a[i].material}).font=font2;
    }
    else if (a[i].bimoc === '광산품') {
      var oi = a[i].material * json[i].수량;
      worksheet.addRow({number: i+2,code: json[i].품목코드, name: json[i].품명,size: json[i].규격,unit: json[i].단위,count: json[i].수량,materialwon: oi, material: a[i].material, labor: a[i].labor, public: a[i].public, industry: a[i].material}).font=font2;
    }
    else if (a[i].bimoc === '농림수산품') {
      var oi = a[i].public * json[i].수량;
      worksheet.addRow({number: i+2,code: json[i].품목코드, name: json[i].품명,size: json[i].규격,unit: json[i].단위,count: json[i].수량,publicwon: oi, material: a[i].material, labor: a[i].labor, public: a[i].public, mine: a[i].material}).font=font2;
    }
    else if (a[i].bimoc === '공사노임') {
      worksheet.addRow({number: i+2,code: json[i].품목코드, name: json[i].품명,size: json[i].규격,unit: json[i].단위,count: json[i].수량, material: a[i].material, labor: a[i].labor, public: a[i].public, noim: a[i].labor}).font=font2;
    }
    else if (a[i].bimoc === '국산') {
      var oi = a[i].public * json[i].수량;
      worksheet.addRow({number: i+2,code: json[i].품목코드, name: json[i].품명,size: json[i].규격,unit: json[i].단위,count: json[i].수량,publicwon: oi, material: a[i].material, labor: a[i].labor, public: a[i].public, domestic: a[i].public}).font=font2;
    }
    else if (a[i].bimoc === '외산') {
      var oi = a[i].public * json[i].수량;
      worksheet.addRow({number: i+2,code: json[i].품목코드, name: json[i].품명,size: json[i].규격,unit: json[i].단위,count: json[i].수량,publicwon: oi, material: a[i].material, labor: a[i].labor, public: a[i].public, foreign: a[i].public}).font=font2;
    }
    else if (a[i].bimoc === '전력수도,도시가스') {
      var oi = a[i].public * json[i].수량;
      worksheet.addRow({number: i+2,code: json[i].품목코드, name: json[i].품명,size: json[i].규격,unit: json[i].단위,count: json[i].수량,publicwon: oi, material: a[i].material, labor: a[i].labor, public: a[i].public, gas: a[i].public}).font=font2;
    }
    var celle = 'E', cellf = 'F', cellg = 'G',cellh ='H',celli ='I',cellj ='J',cellk ='K',celll ='L',cellm ='M',celln ='N',cello ='O',cellp ='P',cellq ='Q',cellr ='R',cells ='S',cellt ='T',cellu ='U';
    var cellv = 'V', cellw = 'W', cellx = 'X', celly ='Y', cella = 'A', cellb = 'B', cellc = 'C', celld = 'D';
    var u9 = i+7;
    var u8 = i+7;
    var text = cella.concat(u8);
    worksheet.getCell(text).fill=fill1;
    worksheet.getCell(text).alignment = { vertical: 'middle', horizontal: 'center'};
    var text = cellb.concat(u8);
    worksheet.getCell(text).fill=fill1;
    var text = cellc.concat(u8);
    worksheet.getCell(text).fill=fill1;
    var text = celle.concat(u9);
    worksheet.getCell(text).border=border1;
    var text = cellf.concat(u9);
    worksheet.getCell(text).border=border1;
    var text = cellg.concat(u9);
    worksheet.getCell(text).border=border4;
    var text = cellh.concat(u9);
    worksheet.getCell(text).border=border1;
    var text = celli.concat(u9);
    worksheet.getCell(text).border=border1;
    var text = cellj.concat(u9);
    worksheet.getCell(text).border=border4;
    var text = cellk.concat(u9);
    worksheet.getCell(text).border=border1;
    var text = celll.concat(u9);
    worksheet.getCell(text).border=border1;
    var text = cellm.concat(u9);
    worksheet.getCell(text).border=border1;
    var text = celln.concat(u9);
    worksheet.getCell(text).border=border4;
    var text = cello.concat(u9);
    worksheet.getCell(text).border=border1;
    var text = cellp.concat(u9);
    worksheet.getCell(text).border=border1;
    var text = cellq.concat(u9);
    worksheet.getCell(text).border=border1;
    var text = cellr.concat(u9);
    worksheet.getCell(text).border=border1;
    var text = cells.concat(u9);
    worksheet.getCell(text).border=border1;
    var text = cellt.concat(u9);
    worksheet.getCell(text).border=border1;
    var text = cellu.concat(u9);
    worksheet.getCell(text).border=border1;
    var text = cellv.concat(u9);
    worksheet.getCell(text).border=border1;
    var text = cellw.concat(u9);
    worksheet.getCell(text).border=border1;
    var text = cellx.concat(u9);
    worksheet.getCell(text).border=border1;
    var text = celly.concat(u9);
    worksheet.getCell(text).border=border1;
    var text = cellb.concat(u9);
    worksheet.getCell(text).border=border2;
    var text = cellc.concat(u9);
    worksheet.getCell(text).border=border3;
  }
  for(var j=2; j<=6; j++){
    var texta = 'A',textb = 'B',textc = 'C';
    var text1 = j;
    var final = texta.concat(text1);
    worksheet.getCell(final).fill=fill1;
    var final = textb.concat(text1);
    worksheet.getCell(final).fill=fill1;
    var final = textc.concat(text1);
    worksheet.getCell(final).fill=fill1;
  }
  worksheet.getCell('A6').alignment = { vertical: 'middle', horizontal: 'center'};
  worksheet.getRow(6).border = {
    top:{style:'thin'},
    right:{style:'dotted'},
    left:{style:'dotted'},
    bottom:{style:'dotted'}
  }
  worksheet.getCell('A2').border ={}
  worksheet.getCell('B2').border ={}
  worksheet.getCell('A6').border ={
    right: {style:'thin'}
  }
  worksheet.getCell('B6').border ={
    top: {style:'thin'},
    left: {style:'thin'},
    right: {style:'thin'},
    bottom:{style:'dotted'}
  }
  worksheet.getCell('C6').border ={
    top: {style:'thin'},
    left: {style:'thin'},
    right: {style:'thin'},
    bottom:{style:'dotted'}
  }

  worksheet.getCell('B3').border = {
    top: {style:'thin'},
    left: {style:'thin'},
    right: {style:'thin'}
  };
  worksheet.getCell('B4').border = {
    left: {style:'thin'},
    right: {style:'thin'}
  };
  worksheet.getCell('B5').border = {
    left: {style:'thin'},
    right: {style:'thin'},
    bottom: {style:'thin'}
  };
  worksheet.getCell('C3').border = {
    top: {style:'thin'},
    left: {style:'thin'},
    right: {style:'thick', color : {argb : 'FF0000FF'}}
  };
  worksheet.getCell('C4').border = {
    left: {style:'thin'},
    right: {style:'thick', color : {argb : 'FF0000FF'}}
  };
  worksheet.getCell('C5').border = {
    left: {style:'thin'},
    right: {style:'thick', color : {argb : 'FF0000FF'}},
    bottom: {style: 'thin'}
  };
  worksheet.getCell('D6').border = {
    top:{style : 'thin'},
    left: {style:'thick', color : {argb : 'FF0000FF'}},
    right: {style:'dotted'},
    bottom: {style:'dotted'}
  };
  worksheet.getCell('C2').border = {
    right:{style:'thick', color : {argb : 'FF0000FF'}}
  };







  writeworkbook.xlsx.writeFile('final9.xlsx');
  console.log("DONE");
}).catch((err)=> {
  console.log(err);
});

res.render('file');
})

app.get("/testa", (req, res)=>{
  res.render("testa")
})

app.get('/edit_db', function (req,res) {
  var id = "a"
  var selected_excel = 'companyjung';



})

app.listen(3000, function(){
  console.log('connected, 3000port');
})

// for(j=0;j<length;j++){
//    worksheet.addRow({item: "oceanfog", name: '경록', contact: "010-3588-6265"});
// }
