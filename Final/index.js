(async() => {
    let workbook = XLSX.read(await (await fetch("./Final.xlsx")).arrayBuffer());
    
    let worksheet = workbook.SheetNames;
   // let excelRowsObjArr = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[worksheet[1]]);
   
    sheet =  workbook.Sheets[worksheet[1]];
    const range = XLSX.utils.decode_range(sheet["!ref"]||"A1");
   
    let i= (range['e'].r) + 1;
    
    while ( i-- ) {
    if((sheet['C'+(i)].v)>(sheet['C'+(i+1)].v)){
        max = (sheet['F'+(i)].v);
    }
 if(c===(sheet['C'+i].v)){
document.getElementById('ahat').innerHTML=sheet['E'+i].v;
document.getElementById('qulty').innerHTML= sheet['F'+i].v+"% Quality <br> vs<br> Benchmark of "+max+"%";
    
 }
}
})()
let d = document.getElementById('qulty').textContent;
console.log(d);