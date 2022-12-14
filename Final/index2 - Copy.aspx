<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link href="./bootstrap.min.css" rel="stylesheet">
    <link href="./docs.css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="./index.css"  />
    <title>Dashboard</title>
    <script src="./bootstrap.bundle.min.js"></script>
    <script src="./chart.js"></script>
    <script src="./xlsx.full.min.js" integrity="sha512-r22gChDnGvBylk90+2e/ycr3RVrDi8DIOkIGNhJlKfuyQM4tIRAI062MaV8sfjQKYVGjOBaZBOA87z+IhZE9DA==" crossorigin="anonymous" referrerpolicy="no-referrer"> //01000000 01100001 01101110 01101011 01101001 01110100 01100100 01100001 </script> 
    <script src="./index.js"></script>
</head>
<body>



<script>




</script>


    <nav id = "navbar-example2" class="navbar fixed-top navbar-expand-lg bg-light px-3" >
        <div class="container-fluid">
          <a class="navbar-brand" href="#">Team Kamal</a>
          <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navbarText" aria-controls="navbarText" aria-expanded="false" aria-label="Toggle navigation">
            <span class="navbar-toggler-icon"></span>
          </button>
          <div class="collapse navbar-collapse" id="navbarText">
            <ul class="nav nav-pills me-auto mb-2 mb-lg-0">
              <li class="nav-item">
                <a class="nav-link active" aria-current="page" href="#AHT">AHT</a>
              </li>
              <li class="nav-item">
                <a class="nav-link" href="#Queues">Queues</a>
              </li>
              <li class="nav-item">
                <a class="nav-link" href="#Benchmark">Benchmark</a>
              </li>
              <li class="nav-item">
                <a class="nav-link" data-bs-toggle="offcanvas" href="#offcanvasExample" role="button" aria-controls="offcanvasExample">
                  Useful Links
                </a>


              </li>
            </ul>

       
            <span class="navbar-text" style="padding-right:10px;">
                <a >Welcome <a id="upi"> _spPageContextInfo.userId 
                  <script>
                  
                  var userID=_spPageContextInfo.userLoginName;
                  
                 document.getElementById("upi").innerHTML = userID ;
                </script></a></a>
               
                <button type="button" class="btn btn-primary">
                    Notifications <span class="badge text-bg-secondary">4</span>
                  </button>
            </span>
          </div>
        </div>
      </nav>


      <div class="offcanvas offcanvas-start" tabindex="-1" id="offcanvasExample" aria-labelledby="offcanvasExampleLabel">
        <div class="offcanvas-header">
          <h5 class="offcanvas-title" id="offcanvasExampleLabel">Useful Links</h5>
          <button type="button" class="btn-close" data-bs-dismiss="offcanvas" aria-label="Close"></button>
        </div>
        <div class="offcanvas-body">
          <div>
          
            
            <div class="list-group">
            
              <a href="#" class="list-group-item list-group-item-action">En_GB</a>
              <a href="#" class="list-group-item list-group-item-action">En_US</a>
              <a href="#" class="list-group-item list-group-item-action">Core</a>
              <a class="list-group-item list-group-item-action disabled">A disabled link item</a>
            </div>



          </div>
          <div class="dropdown mt-3">
            <button class="btn btn-secondary dropdown-toggle"  type="button" data-bs-toggle="dropdown">
              Stuff
            </button>
            <ul class="dropdown-menu"  id = "dropupown" >

           
              <li   > <a class="dropdown-item"  href="#">Immediate Links like of quip fetches <br> from a list of excel created</a></li>
            </ul>
          </div>
        </div>
      </div>
      <div data-bs-spy="scroll" data-bs-target="#navbar-example2" data-bs-offset="0"  tabindex="0">

    <div id ="AHT" class="container-fluid text-center  m-0 border-0 bd-example bd-example-row" style="padding-top: 60px;">
        <div class="row">
          <div class="col">
            <div class="card text-dark bg-light mb-3" >
            <div class="card-header">AHT</div>
            <div class="card-body">
                <br>
                
              <h5 class="card-title p-1" id = "ahat">AHT Score here</h5>
             
             
             
              <br>
              <p class="card-text ">
                
                <canvas class ="w-100" id="myChartfill" >
                    <script>

console.log("Test"); // rough i know but will centralise later this is ifor demo
(async() => {
    let workbook = XLSX.read(await (await fetch("./Final.xlsx")).arrayBuffer());
    
    let worksheet = workbook.SheetNames;
   // let excelRowsObjArr = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[worksheet[1]]);
   
    let sheet =  workbook.Sheets[worksheet[1]];
     
    const range = XLSX.utils.decode_range(sheet["!ref"]||"A1");
    console.log((range['e'].r) + 1);
    console.log(c1);
    let i= (range['e'].r) + 1;
    var z=0;
let    arrayforchartquality = [];
    while ( i-- ) {
      console.log(sheet['C'+i].v);
 if(c1===(sheet['C'+i].v)){
  console.log(sheet['C'+i].v);
  console.log("found"+worksheet);
  //let qulty= sheet['F'+i].v;
  //let Ahat=sheet['E'+i].v;
  worksheet.forEach(element => {
    sheet1 =  workbook.Sheets[element];
    console.log(sheet1['F'+i].v);
    arrayforchartquality[z]= sheet1['F'+i].v;
    
    z++;
  });
break;
 }

}
console.log(arrayforchartquality);
                        chartData4 = 
                        {
                            labels: worksheet,
                        datasets: [{
                        label: 'My First Dataset',
                        data: arrayforchartquality,
                        fill: true,
                        borderWidth: 1.0,
                        backgroundColor: 'rgba(255, 99, 132, 0.2)',
                        pointStyle : 'line',
                        
                     
                        }],
                        }
                        ;
                        const myChartfill = document.getElementById('myChartfill');
                        
                        if(myChartfill!=null){
                        new Chart(myChartfill, {
                           
                          type: 'line',
                        data: chartData4,
                       
                        options: { 
                            events: [],
                            plugins: { tooltip: { enabled: false} ,
                            legend: {
                            display: false},
                            
                                    },
                                    scales: {
                                         y: {
                                            ticks: {
                                             display: false,
                                          },
                                          grid: {
                                          display: false,
                                          },
                                        },
                                         x: {
                                            ticks: {
                                             display: false,
                                          },
                                          grid: {
                                          display: false,
                                          },
                                          },
                                            },

                        },
                    }
                        )}
                  }
)()
                      
                      </script>


                </canvas>

              </p>
            </div>
          </div>
        </div>
          <div class="col">
            <div class="card text-dark bg-light mb-3" >
                <div class="card-header">Quality</div>
                <div class="card-body">
                    <br>
             
              
                  <h5 class="card-title p-1" id = "qulty">Quality Score here</h5>
                 
              
              <br>
                  <p class="card-text">
                    <canvas class ="w-100" id="myChartfill2" >
                        <script>


let c="stu";
console.log("Test"); // rough i know but will centralise later this is ifor demo
(async() => {
    let workbook = XLSX.read(await (await fetch("./Final.xlsx")).arrayBuffer());
    
    let worksheet = workbook.SheetNames;
   // let excelRowsObjArr = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[worksheet[1]]);
   
    sheet =  workbook.Sheets[worksheet[1]];
   
    const range = XLSX.utils.decode_range(sheet["!ref"]||"A1");
    console.log((range['e'].r) + 1);
    console.log(c);
    let i= (range['e'].r) + 1;
    var z=0
let    arrayforchartAHT = [];
    while ( i-- ) {
      console.log(sheet['C'+i].v);
 if(c===(sheet['C'+i].v)){
  console.log(sheet['C'+i].v);
  console.log("found");
  //let qulty= sheet['F'+i].v;
  //let Ahat=sheet['E'+i].v;
  worksheet.forEach(element => {
    sheet1 =  workbook.Sheets[element];
    arrayforchartAHT[z]= sheet1['E'+i].v;
    z++;
  });
break;
 }

}


                            chartData5 = 
                            {
                           
                                labels: worksheet,
                            datasets: [{
                            label: 'My First Dataset',
                            data: arrayforchartAHT,
                            fill: true,
                            borderWidth: 1.0,
                            backgroundColor: 'rgba(255, 99, 132, 0.2)',
                            pointStyle : 'line',
                            
                         
                            }],
                            }
                            ;
                            const myChartfill2 = document.getElementById('myChartfill2');
                            
                            if(myChartfill2!=null){
                            new Chart(myChartfill2, {
                               
                              type: 'line',
                            data: chartData5,
                           
                            options: { 
                                events: [],
                                plugins: { tooltip: { enabled: false} ,
                                legend: {
                                display: false},
                                
                                        },
                                        scales: {
                                             y: {
                                                ticks: {
                                                 display: false,
                                              },
                                              grid: {
                                              display: false,
                                              },
                                            },
                                             x: {
                                                ticks: {
                                                 display: false,
                                              },
                                              grid: {
                                              display: false,
                                              },
                                              },
                                                },
    
                            },
                        }
                            )}
                          
                  }
)()
                          
                          </script>
    
    
                    </canvas>
    
                  </p>


                  
                </div>
              </div>

          </div>
          <div class="col">
            <div class="card text-dark bg-light mb-3" >
                <div class="card-header">Production Hour</div>
                <div class="card-body">
                    <br>
                    
                  <h5 class="card-title p-1">Production Hour Through API</h5>
                  
                  <br>
                  <br>
                  <p class="card-text">
                    <canvas class ="w-100" id="myChartbar2" >
                        <script>
                            chartData6 = 
                            {
                           
                                labels: ['Jan', 'Feb', 'Mar', 'Apr'],
                            datasets: [{
                            label: 'My First Dataset',
                            data: [10, 20, 30, 40],
                                 borderColor: 'rgb(255, 99, 132)',
                              backgroundColor: 'rgba(255, 99, 132, 0.2)',
                            backgroundColor: 'rgba(255, 99, 132, 0.2)',
                            
                            
                         
                            }],
                            }
                            ;
                            const myChartbar2 = document.getElementById('myChartbar2');
                            
                            if(myChartbar2!=null){
                            new Chart(myChartbar2, {
                               
                              type: 'bar',
                            data: chartData6,
                           
                            options: { 
                                events: [],
                                plugins: { tooltip: { enabled: false} ,
                                legend: {
                                display: false},
                                
                                        },
                                        scales: {
                                             y: {
                                                ticks: {
                                                 display: false,
                                              },
                                              grid: {
                                              display: false,
                                              },
                                            },
                                             x: {
                                                ticks: {
                                                 display: false,
                                              },
                                              grid: {
                                              display: false,
                                              },
                                              },
                                                },
    
                            },
                        }
                            )}
                          
                          
                          </script>
    
    
                    </canvas>
                  </p>
                </div>
              </div>

          </div>
          <div class="col" style=" height:inherit;" >
            <div class="card">
                <div class="card-body">
                    <canvas id="myChart" >
                        <script> 
                        
                        

//let                      c="stu";
console.log("Test"); // rough i know but will centralise later this is ifor demo
(async() => {
    let workbook = XLSX.read(await (await fetch("./Final.xlsx")).arrayBuffer());
    
    let worksheet = workbook.SheetNames;
    
   // let excelRowsObjArr = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[worksheet[1]]);
   
    sheet =  workbook.Sheets[worksheet[1]];
    const range = XLSX.utils.decode_range(sheet["!ref"]||"A1");
    console.log((range['e'].r) + 1);
    console.log(c);
    let i= (range['e'].r) + 1;
    var z=0
let    arrayforchartpi = [];
let arrayforchartpi2 = [];
    while ( i-- ) {
      console.log(sheet['C'+i].v);
 if(c1===(sheet['C'+i].v)){
  console.log(sheet['C'+i].v);
  console.log("found");
  //let qulty= sheet['F'+i].v;
  //let Ahat=sheet['E'+i].v;

    arrayforchartpi[0]= sheet['E'+i].v;

    arrayforchartpi2[0]= sheet['E1'].v;
    arrayforchartpi[1]= sheet['F'+i].v;
    arrayforchartpi2[1]= sheet['F1'].v;
    arrayforchartpi[2]= sheet['H'+i].v;
    arrayforchartpi2[2]= sheet['H1'].v;
    arrayforchartpi[3]= sheet['J'+i].v;
    arrayforchartpi2[3]= sheet['J1'].v;
    arrayforchartpi[4]= sheet['K'+i].v;
    arrayforchartpi2[4]= sheet['K1'].v;
break;
 }

}


                        
                        chartData = {
                            labels: arrayforchartpi2,
                            datasets: [{
                              label: 'My First Dataset',
                              data: arrayforchartpi,
                              backgroundColor: [
                                'rgb(255, 99, 132)',
                                'rgb(54, 162, 235)',
                                'rgb(255, 205, 86)',
                                'rgb(255, 99, 132)'
                              ],
                              hoverOffset: 4
                            }]
                          };
                              
                    
                          const myChart = document.getElementById('myChart');
                    
                    
                    
                    if(myChart){
                      new Chart(myChart, {
                        
                          type: 'doughnut',
                          
                      data: chartData,
                      options: {
                          scales: {
                              y:  {
                                                ticks: {
                                                 display: false,
                                              },
                                             
                                            },
                          }
                      }
                      }
                      )};
                         
                  }
)()
                          
                      </script>

                    </canvas>
                </div>
              </div>

          </div>
        </div>
        </div>
        <div id = "Queues" class="container-fluid text-center p-3 m-0 border-0 bd-example bd-example-row">
        <div class="row">
            <div class="col-8 container-fluid ">
                <div class="card " >
                    <div class="card-body">
               
                <table class="table table-hover" >
          
            
            </table>
            <script>
              /*
                let table = document.querySelector(".table");
                (
                    async() => {
                        let workbook = XLSX.read(await (await fetch("./Book.xlsx")).arrayBuffer());
                      
                        let worksheet = workbook.SheetNames;
                        console.log(worksheet);
                        worksheet.forEach(name => {
                            let html = XLSX.utils.sheet_to_html(workbook.Sheets[name]);
                            table.innerHTML += `
                            <h3>${name}</h3>${html}
                            `;
                            
                        })
                    }
                )()
                */
            </script>
            </div>
        </div>
            </div>
            <div class="col-4  ">
                <div class="card">
                    <div class="card-body">

                <canvas class = "w-100" id="myChart2" >
                    <script>





                    

//let                      c="stu";
console.log("Test"); // rough i know but will centralise later this is ifor demo
(async() => {
    let workbook = XLSX.read(await (await fetch("./Final.xlsx")).arrayBuffer());
    
    let worksheet = workbook.SheetNames;
    
   // let excelRowsObjArr = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[worksheet[1]]);
   
    
    const range = XLSX.utils.decode_range(sheet["!ref"]||"A1");
    console.log((range['e'].r) + 1);
    console.log(c1);
    var i= (range['e'].r) + 1;
    var z=0;
    var countersheet=0;
let    arrayforaht2 = [];
let arrayforquality2 = [];
let arrayforshrinkage2 = [];
let arrayforcomplian2 = [];
worksheet.forEach(element => {
  console.log(element);

  sheet =  workbook.Sheets[[element]];
  console.log((sheet['C'+i].v));
  i= (range['e'].r) + 1;
    while ( i-- ) {
      console.log((sheet['C'+i].v));
 if(c1===(sheet['C'+i].v)){
 
  //let qulty= sheet['F'+i].v;
  //let Ahat=sheet['E'+i].v;

    arrayforaht2[countersheet]= sheet['E'+i].v;
    arrayforquality2[countersheet]= sheet['F'+i].v;
    arrayforshrinkage2[countersheet]= sheet['K'+i].v;
    arrayforcomplian2[countersheet]= sheet['J'+i].v;
    break;
 }

}
console.log(countersheet);
countersheet++;
});



                        chartData1 = 
                        {
                            labels: worksheet,
                        datasets: [{
                        label: 'AHT',
                        data: arrayforaht2,
                        fill: false,
                        borderColor: 'rgb(75, 192, 192)',
                        tension: 0.1,
                        hoverOffset: 4
                        },
                      {
                        label: 'Quality',
                        data: arrayforquality2,
                        fill: false,
                        borderColor: 'rgb(54, 162, 235)',
                        tension: 0.1,
                        hoverOffset: 4
                        },
                        
                        {
                            
                        label: 'Shrinkage',
                        data: arrayforshrinkage2,
                        fill: false,
                        borderColor: 'rgb(255, 205, 86)',
                        tension: 0.1,
                        hoverOffset: 4
                        },
                        {
                        label: 'Compliance',
                        data: arrayforcomplian2,
                        fill: false,
                        borderColor: 'rgb(75, 192, 192)',
                        tension: 0.1,
                        hoverOffset: 4
                        }]
                        };
                        const myChart2 = document.getElementById('myChart2');
                        
                        if(myChart2!=null){
                        new Chart(myChart2, {
                          type: 'line',
                        data: chartData1,
                        options: {
                            scales: {
                                y: {
                                    beginAtZero: true
                                }
                            }
                        }
                        }
                        )}
                        
                  }
)()
                      
                      </script>


                </canvas>
</div>
</div>

            </div>
          </div>
        </div>
        <div id = "Benchmark" class="container-fluid text-center p-3 m-0 border-0 bd-example bd-example-row">
        <div class="row">
            <div class="col-sm">Chart to show team's Live AHT cloaking the name of everyone else exept _spPageContextInfo.userId</div>
            <div class="col-sm">

                <div class="card  ">
                    <div class="card-body h-100">
                <canvas id="myChartlinewithbar22" class="w-100">
                <script>
                 const chartData11 = {
                    labels: [
                      'January',
                      'February',
                      'March',
                      'April'
                    ],
                    datasets: [{
                      type: 'bar',
                      label: "Team's Benchmark",
                      data: [10, 20, 30, 40],
                      borderColor: 'rgb(255, 99, 132)',
                      backgroundColor: 'rgba(255, 99, 132, 0.2)'
                    }, {
                      type: 'line',
                      label: 'Your Quality',
                      data: [30, 90, 60, 9],
                      fill: false,
                      borderColor: 'rgb(54, 162, 235)'
                    }]
                  };
                        
                  
                    const myChartlinewithbar22 = document.getElementById('myChartlinewithbar22');
                  
                  
                  
                  if(myChartlinewithbar22){
                  new Chart(myChartlinewithbar22, {
                    type: 'scatter',
                    data: chartData11,
                    options: {
                      scales: {
                        y: {
                          beginAtZero: true
                        }
                      }
                    }
                  }
                  )}
                  ;</script>

</canvas>
</div>
</div>

            </div>
            <div class="col-sm">
                
                <div class="card">
                    <div class="card-body">

<canvas id="myChartlinewithbar" >
<script> 











const chartData3 = {
  labels: [
    'January',
    'February',
    'March',
    'April'
  ],
  datasets: [{
    type: 'bar',
    label: "Team's Benchmark",
    data: [10, 20, 30, 40],
    borderColor: 'rgb(255, 99, 132)',
    backgroundColor: 'rgba(255, 99, 132, 0.2)'
  }, {
    type: 'line',
    label: 'Your AHT',
    data: [30, 50, 21, 9],
    fill: false,
    borderColor: 'rgb(54, 162, 235)'
  }]
};
      

  const myChartlinewithbar = document.getElementById('myChartlinewithbar');



if(myChartlinewithbar){
new Chart(myChartlinewithbar, {
  type: 'scatter',
  data: chartData3,
  options: {
    scales: {
      y: {
        beginAtZero: true
      }
    }
  }
}
)};


</script>

</canvas>
</div>
</div>



            </div>
          </div>
        </div>
  </div>
 
    <button class = "float" id = "csstext" type="button" data-bs-toggle="modal"  data-bs-target="#staticBackdrop"   > Click

    </button>
 
   

   
    
    <!-- Modal -->
    <div class="modal fade" id="staticBackdrop" data-bs-backdrop="static" data-bs-keyboard="false" tabindex="-1" aria-labelledby="staticBackdropLabel" aria-hidden="true">
      <div class="modal-dialog">
        <div class="modal-content">
          <div class="modal-header">
            <h5 class="modal-title" id="staticBackdropLabel">Monthly Feedback</h5>
            <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
          </div>
          <div class="modal-body">
            <div class="mb-3">
              <label for="exampleFormControlTextarea1" class="form-label">DAs Input</label>
              <textarea class="form-control" id="exampleFormControlTextarea1" rows="3"></textarea>
            </div>
            <div class="mb-3">
              <label for="exampleFormControlTextarea1" class="form-label">Manager's Input</label>
              <textarea class="form-control" id="exampleFormControlTextarea1" rows="3"></textarea>
            </div>
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
            <button type="button" class="btn btn-primary">Save</button>
          </div>
        </div>
      </div>
    </div>
 
</body>
</html>