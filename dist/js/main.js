function html(Obj) {
  Object.keys(Obj).forEach((a, b) =>  {
    document.getElementById(a).innerHTML = Object.values(Obj)[b];
  });
}
localStorage.removeItem('ExcelData');
const Toast = Swal.mixin({
  toast: true,
  position: 'top-end',
  showConfirmButton: false,
  timer: 3000,
  timerProgressBar: true,
  onOpen: (toast) => {
    toast.addEventListener('mouseenter', Swal.stopTimer)
    toast.addEventListener('mouseleave', Swal.resumeTimer)
  }
})

let fileInput = document.getElementById("fileUploader"),
    dragFile = document.querySelector('.upload'),
    previewData = document.getElementById('seeData'),
    createFile = document.getElementById('xlxsFile'),
    saveFile = document.getElementById('saveFile');
function handleFile(e) {
  Toast.fire({
    icon: 'info',
    title: 'Proses...'
  });
  var files = e.target.files, f = files[0];
  var reader = new FileReader();
  reader.onload = function(e) {
    var data = new Uint8Array(e.target.result);
    var workbook = XLSX.read(data, {type: 'array'});

    workbook.SheetNames.forEach((name)=>  {
      let rowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[name]),
          paperContent = '',
          textContent = '',
          dataArrayContent = [],
          lengthRowObject = Math.floor(rowObject.length/8);
      localStorage.setItem('ExcelData', JSON.stringify(rowObject));
      rowObject.forEach((values, keys) =>  {
        textContent += '<tr><td>' + values.No +'</td><td>' + values.no_ujian +'</td><td>'+ values.nama +'</td><td>'+ values.tanggal_lahir + '</td><td>'+ values.Ruang +'</td></tr>';
        paperContent += 
        `
        <div class="table">
          <table border="0" width="100%">
            <tr style="text-transform: uppercase;">
              <th width="100"><img src="`+ dataList.logo +`" width="60px" height="60px" alt="logos"></th>
              <th>
                <strong>
                    <p>`+ dataList.title +`</p>
                    <p>`+ dataList.subtitle +`</p>
                </strong>
                <small>`+ dataList.alamat +`</small>
              </th>
            </tr>
            <tr>
                <td colspan="2">
                <hr>
                  <p style="text-align: center;font-weight: bold; text-transform: uppercase;">`+ dataList.cardname +`</p>
                </td>
            </tr>
            <tr>
              <table border="0" style="width:100%;text-align:left;padding-left:5px;">
                <tr> 
                  <td> Nomer Ujian </td>
                  <td>: ${values.no_ujian} </td>
                </tr>
                <tr> 
                  <td> Nama </td>
                  <td>: ${values.nama} </td>
                </tr>
                <tr> 
                  <td> Tanggal Lahir </td>
                  <td>: ${values.tanggal_lahir} </td>
                </tr>
                <tr> 
                  <td> Ruang </td>
                  <td>: ${values.Ruang} </td>
                </tr>
                <tr style="text-align:right;"> 
                  <td>  </td>
                  <td colspan="2">${dataList.date} </td>
                </tr>
                <tr> 
                  <td>  </td>
                  <td style="text-align:center;">${dataList.ttd.title} </td>
                  <td>  </td>
                </tr>
                <tr> 
                  <td>  </td>
                  <td></td>
                  <td>  </td>
                </tr>
                <tr> 
                  <td>  </td>
                  <td></td>
                  <td>  </td>
                </tr>
                <tr> 
                  <td>  </td>
                  <td></td>
                  <td>  </td>
                </tr>
                <tr> 
                  <td>  </td>
                  <td></td>
                  <td>  </td>
                </tr>
                <tr> 
                  <td>  </td>
                  <td></td>
                  <td>  </td>
                </tr>
                <tr> 
                  <td>  </td>
                  <td></td>
                  <td>  </td>
                </tr>
                <tr> 
                  <td>  </td>
                  <td></td>
                  <td>  </td>
                </tr>
                <tr> 
                  <td>  </td>
                  <td></td>
                  <td>  </td>
                </tr>
                <tr> 
                  <td>  </td>
                  <td></td>
                  <td>  </td>
                </tr>
                <tr> 
                  <td>  </td>
                  <td></td>
                  <td>  </td>
                </tr>
                <tr> 
                  <td>  </td>
                  <td style="text-align:center;">${dataList.ttd.name} </td>
                  <td>  </td>
                </tr>
                <tr> 
                  <td>  </td>
                  <td style="text-align:center;">NIP :  ${dataList.ttd.nip} </td>
                  <td>  </td>
                </tr>
              </table>
            </tr>
          </table>
        </div>`;
        if((keys+1)%8 == 0) {
            dataArrayContent.push(paperContent);
            paperContent = '';
        } else if((keys/8) >= lengthRowObject) {
            if((keys+1) == rowObject.length)  {
              dataArrayContent.push(paperContent);
              paperContent = '';
            }
        }  
      });
      let div = '';
      dataArrayContent.forEach((d)=>{
        div +=  `<div class="pageBreak row" id="resultPage">${d}</div>`;
      })
      html({
        page : div
      });
    });

    
    
    Toast.fire({
      icon: 'success',
      title: 'Import Berhasil'
    });
    document.querySelector('.upload').style.display = 'none';
    document.querySelector('.container').style =  "position:fixed;";
  };
  reader.readAsArrayBuffer(f);
}
fileInput.addEventListener('change', handleFile, false);
dragFile.addEventListener('drop', handleFile, false);

previewData.addEventListener('click', (e) => {
  let data = JSON.parse(localStorage.getItem('ExcelData')),
      textContent = '';
  if(data == null) {
    return swal.fire({
      title : "Oppss...",
      text : "Data tidak ditemukan",
      icon : 'error'
    });
  }
  data.forEach((value)=> {
    textContent += `
      <tr>
        <td> ${value.No} </td>
        <td> ${value.no_ujian} </td>
        <td> ${value.nama} </td>
        <td> ${value.tanggal_lahir} </td>
        <td> ${value.Ruang} </td>
      </tr>
    `;
  })
  Swal.fire({
    title : 'Data List',
    width : 800,
    html : `
    <table border="1" style="width: 100%;display: block; border:#000 2px;">
      <thead>
          <td> No </td>
          <td> Nomer Ujian </td>
          <td> Nama </td>
          <td> Tanggal Lahir </td>
          <td> Ruang </td>
      </thead>
      <tbody>
        ${textContent}
      </tbody>
      <tfoot>
          Jumlah Total : <strong> ${data.length} </strong> siswa
      </tfoot>
    </table>`
  });
})

createFile.addEventListener('click', (e) => {
  var wb = XLSX.utils.book_new();
  var ws_name = "SheetJS";

  /* make worksheet */
  var ws_data = [
    [ "No", "no_ujian", "nama", "tanggal_lahir", "Ruang"],
    [  1 ,  '0021-123-202' ,  'Your Name' ,  "-" ,  5 ],
    [  2 ,  '0021-123-203' ,  'Your Name HUH' ,  "-" ,  5 ],
  ];
  var ws = XLSX.utils.aoa_to_sheet(ws_data);

  /* Add the worksheet to the workbook */
  XLSX.utils.book_append_sheet(wb, ws, ws_name);
  /* bookType can be any supported output type */
  var wopts = { bookType:'xlsx', bookSST:false, type:'array' };

  var wbout = XLSX.write(wb,wopts);

  /* the saveAs call downloads a file on the local machine */
  saveAs(new Blob([wbout],{type:"application/octet-stream"}), "template.xlsx");

})

saveFile.addEventListener('click', (e) =>  {
    window.print();
})
