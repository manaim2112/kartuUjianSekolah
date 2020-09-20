$(document).ready(function(){

    let exelData = {};
    let datalist = {
      "title"     : "PEMERINTAH KABUPATEN PASURUAN <br> KEMENTERIAN AGAMA",
      "subtitle"  : "Mts Sunan Ampel",
      "alamat"    : "KarangAnyar Kraton Pasuruan",
      "logo"      : "./dist/images/logo.png",
      "cardname"  : "kartu Ujian Tengah Semester Genap",
      "date"      : "Pasuruan, 09 Maret 2020",
      "ttd"       : {
          "title"     : "Kepala Sekolah",
          "name"      : "IKHWAN, S.Ag",
          "nip"       : ""
      }
    };
    
    document.getElementById('title').innerHTML = datalist.title;
    document.getElementById('subtitle').innerHTML = datalist.subtitle;
    document.getElementById('alamat').innerHTML = datalist.alamat;
    document.getElementById('logo').innerHTML = '<img src="'+ datalist.logo +'" width="60px" height="60px" alt="logos">';
    document.getElementById('cardname').innerHTML = datalist.cardname;
    document.getElementById('date').innerHTML = datalist.date;
    document.getElementById('ttd-title').innerHTML = datalist.ttd.title;
    document.getElementById('ttd-name').innerHTML = datalist.ttd.name;
    document.getElementById('ttd-nip').innerHTML = datalist.ttd.nip;
    
  $("#fileUploader").change(function(evt){
        var selectedFile = evt.target.files[0];
        var reader = new FileReader();
        reader.onload = function(event) {
          var data = event.target.result;
          var workbook = XLSX.read(data, {
              type: 'binary'
          });
          workbook.SheetNames.forEach(function(sheetName) {
            
              var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
              var json_object = JSON.stringify(XL_row_object);
              //document.getElementById("jsonObject").innerHTML = json_object;
            
                exelData = JSON.parse(json_object);
                let textContent = '';
                let paperContent = '';
                
                for(var i=0; i < exelData.length; i++)
                {
                    /**
                     * Datatable List
                     */
                    textContent += '<tr><td>' + exelData[i].No +'</td><td>' + exelData[i].no_ujian +'</td><td>'+ exelData[i].nama +'</td><td>'+ exelData[i].tanggal_lahir + '</td><td>'+ exelData[i].Ruang +'</td></tr>';
                    
                    /**
                     * Data PaperContent
                     */
                    paperContent += `
                        <div class="table">
                          <table border="0">
                            <tr style="text-transform: uppercase;">
                              <th width="100"><img src="`+ datalist.logo +`" width="60px" height="60px" alt="logos"></th>
                              <th>
                                <strong>
                                    <p>`+ datalist.title +`</p>
                                    <p>`+ datalist.subtitle +`</p>
                                </strong>
                                <small>`+ datalist.alamat +`</small>
                              </th>
                            </tr>
                          </table>
                          <hr>
                          <div class="content">
                            <p style="text-align: center;font-weight: bold; text-transform: uppercase;">`+ datalist.cardname +`</p>
                            <div class="row-3 text-left">
                              <div class="col-1">
                                <p>Nomer Ujian</p>
                                <p>Nama</p>
                                <p>Tanggal Lahir</p>
                                <p>Ruang</p>
                              </div>
                              <div class="col-2">
                                <p>: `+ exelData[i].no_ujian +`</p>
                                <p>: `+ exelData[i].nama +`</p>
                                <p>: `+ exelData[i].tanggal_lahir +`</p>
                                <p>: `+ exelData[i].Ruang +`</p>
                                <div class="foot" style="text-align: center;">
                                  <p>`+ datalist.date +`</p>
                                  <p>`+ datalist.ttd.title +`</p>
                                  <br><small>TTD.</small>
                                  <p>`+ datalist.ttd.name +`</p>
                                  <p> NIP : `+ datalist.ttd.nip +`</p>
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>`;
                    // Page Break
                    if((i+1)%8 == 0) {
                      paperContent += '<span id="nextpage"></span>';
                    }

                }
                document.getElementById("dataJson").innerHTML = textContent;
                document.getElementById("lengthJSON").innerHTML = exelData.length;
                document.getElementById("resultPage").innerHTML = paperContent;
            });
        };

        reader.onerror = function(event) {
          console.error("File could not be read! Code " + event.target.error.code);
        };

        reader.readAsBinaryString(selectedFile);
        
  });

  /**
   * Create .xlsx file with data template
   */
  /**
   * Create .xlsx file with data template
   */
  // let tempXLSX = document.querySelector('#downloadtemplate');
  // tempXLSX.addEventListener('click', (event) => {

  //   // Create workbook
  //   let wb = XLSX.utils.book_new();
  //   // Update properties
  //   wb.Props = {
  //     Title : "Aplikasi Kartu Ujian Sekolah",
  //     Subject : "Vol.01",
  //     Author : "@manaim2112",
  //     CreatedDate : new Date(2020,12,21)
  //   };
  //   // SheetName
  //   wb.SheetNames.push("Sheet 1");
  //   // Content Data
  //   let ws_data = [
  //     ['hello', 'world']
  //   ];
  //   let ws = XLSX.utils.aoa_to_sheet(ws_data);
  //   wb.Sheets["Test Sheet!"] = ws;
  //   // download xlsx
  //   let wbout = XLSX.write(wb, {
  //     bookType : 'xlsx',
  //     type : 'binary'
  //   });
  // });

  document.querySelector("#btnPrint").addEventListener('click', (event) => {
    // Get the element.
    var element = document.getElementById('page');

    // Generate the PDF.
    html2pdf().from(element).set({
      margin: 0.2,
      filename: 'download.pdf',
      html2canvas: { scale: 1 },
      image : {type: 'jpeg', quality:0.98},
      jsPDF: {orientation: 'portrait', unit: 'in', format: 'legal', compressPDF: false},
      pagebreak : { mode :'css', before: '#nextpage'}

    }).save();
  });
});
