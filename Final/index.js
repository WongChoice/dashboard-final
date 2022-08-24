let worksheetname = 'January';
let c1="stu";

function secondsToHMS(secs) {
    function z(n) { return (n < 10 ? '0' : '') + n; }
    var sign = secs < 0 ? '-' : '';
    secs = Math.abs(secs);
    return sign + z(secs / 3600 | 0) + ':' + z((secs % 3600) / 60 | 0) + ':' + z(secs % 60);


}
function hmsToSeconds(s) {
    var b = s.toString().split(':');
    return b[0] * 3600 + b[1] * 60 + (+b[2] || 0);
}



(async() => {
    let workbook = XLSX.read(await (await fetch("./Final.xlsx")).arrayBuffer());
    
    let worksheet = workbook.SheetNames;
    var counter = 0;
   // let excelRowsObjArr = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[worksheet[1]]);

    
   
    sheet =  workbook.Sheets[worksheetname];
    const range = XLSX.utils.decode_range(sheet["!ref"]||"A1");
   console.log(sheet['A1'].v);
    let i= (range['e'].r) ;
    console.log(i);
    var   max =0;
    var   max1 =0;
    let  ahtofda;
    let  qualityofda;
    while ( i-- ) {
        if (i>1){
            console.log((sheet['F'+(i)].v));
    if(max<(sheet['F'+(i)].v)){
        max = (sheet['F'+(i)].v);
    }
    if(max<(sheet['E'+(i)].v)){
        max1 = (sheet['E'+(i)].v);
    }

 if(c1===(sheet['C'+(i)].v)){
  ahtofda =sheet['E'+(i)].v;
  qualityofda = sheet['F'+(i)].v;
} 
}
}

document.getElementById('ahat').innerHTML=ahtofda +" AHT <br> vs<br> Benchmark of "+max1 ;
document.getElementById('qulty').innerHTML= qualityofda+"% Quality <br> vs<br> Benchmark of "+max+"%";


})();
(async() => {
    let workbook = XLSX.read(await (await fetch("./Book.xlsx")).arrayBuffer());
    
    let worksheet = workbook.SheetNames;
    var counter = 0;
   // let excelRowsObjArr = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[worksheet[1]]);
  
let    column1=[];
   let column2=[];
   let column3=[];
   let arraytotb=[];
   arraytotb[0] = "<tr><td><b> Date</b> </td><td><b> Queue </b></td><td> <b>Time Processed/Count</b> </td></tr>"
   var nextname = 0;
   worksheet.forEach(sheetKaNaam => {
    const  sheet =  workbook.Sheets[sheetKaNaam];
    console.log(sheet);
   var sumofcolumn=0;
var count =0;
    const range = XLSX.utils.decode_range(sheet["!ref"]||"A1");
  let seconsd;
    let i= (range['e'].r) ;
    console.log(i);
    while ( i-- ) {

    if((sheet['A'+(i+1)].v)===(sheet['A'+(i+2)].v)){
     column1[nextname] = (sheet['A'+(i+1)].v);
     
      seconsd = hmsToSeconds(sheet['B'+(i+1)].v);
      console.log(sheet['B'+(i+1)].v);
      console.log(hmsToSeconds('10:30:29'));
     console.log("hsm"+(sheet['B'+(i+1)].v));
     sumofcolumn = sumofcolumn + seconsd;


      count = count+1;
    }
    else{
        
        column2[nextname]=secondsToHMS(sumofcolumn);
        
column3[nextname] = column2[nextname] +"/"+count;

arraytotb[nextname+1] = "<tr><td>"+column1[nextname]+"</td><td>"+ column2[nextname]+"</td><td>"+ column3[nextname]+"</td></tr>";
nextname = nextname+1;
count=0;
    }
}
});

console.log(arraytotb);
document.querySelector(".table").innerHTML = arraytotb;
})()
