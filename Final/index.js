let worksheetname = 'January';
let c1="stu";
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
    let  ahtofda;
    let  qualityofda;
    while ( i-- ) {
        if (i>1){
            console.log((sheet['F'+(i)].v));
    if(max<(sheet['F'+(i)].v)){
        max = (sheet['F'+(i)].v);
    }

 if(c1===(sheet['C'+(i)].v)){
  ahtofda =sheet['E'+(i)].v;
  qualityofda = sheet['F'+(i)].v;
} 
}
}

document.getElementById('ahat').innerHTML=ahtofda;
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
   
    let i= (range['e'].r) ;
    console.log(i);
    while ( i-- ) {

    if((sheet['A'+(i+1)].v)===(sheet['A'+(i+2)].v)){
     column1[nextname] = (sheet['A'+(i+1)].v);
     sumofcolumn = sumofcolumn + (sheet['B'+(i+1)].v);
      count = count+1;
    }
    else{
        
        column2[nextname]=sumofcolumn;
        
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
