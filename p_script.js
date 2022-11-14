// Method to upload a valid excel file
function upload() {
    var files = document.getElementById('file_upload').files;
    if(files.length==0){
      alert("Please choose any file...");
      return;
    }
    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX') {
      //Here calling another method to read excel file into json
      excelFileToJSON(files[0]);
    }
    else{
      alert("Please select a valid excel file.");
    }
  }
  //Method to read excel file and convert it into JSON 
  function excelFileToJSON(file){
    try {
      var reader = new FileReader();
      reader.readAsBinaryString(file);
      reader.onload = function(e) {
        var data = e.target.result;
        var workbook = XLSX.read(data, {
          type : 'binary'
        }
        );
        var result = {
        };
        var firstSheetName = workbook.SheetNames[0];
        //reading only first sheet data
        var jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName]);
        //displaying the json result into HTML table
        displayJsonToHtmlTable(jsonData);
      }
    }
    catch(e){
      console.error(e);
    }
  }
  //Method to display the data in HTML Table
  function displayJsonToHtmlTable(jsonData){
    // var p=document.getElementById("display_excel_data");
    if(jsonData.length>0){
      var htmlData='<br/>';
      for(var i=0;i<jsonData.length;i++){
        var row=jsonData[i];
        if(row["#"]==undefined || row["Phase"]==undefined || row["Parameter Name"]==undefined || row["Attribute Name"]==undefined || row["Attribute Value"]==undefined)
        continue
        else
        {
          num = new String(row['#']).trim().replace('.','')
          htmlData+=`
          <html>
          <div class="colored" id="demo${i}" onclick=copyToClipboard(selectElementContents(document.getElementById("demo${i}")))>The following screenshot displays the execution of the locked parameter ${row["Parameter Name"]} inside the phase ${row["Phase"]}.<br/>
          <table>
          <tr>
          <th>Expression</th>
          <th>Value</th>
          </tr>
          <tr>
            <td>${row["Attribute Value"]}</td>
            <td></td>
          </tr>
          </table>
          </div>
          </html>`
        }
      }
      // p.innerHTML=htmlData;
      document.getElementById('display_excel_data').innerHTML = htmlData
    }
    else{
      table.innerHTML='There is no data in Excel';
    }
    // document.getElementById('demo').onclick = navigator.clipboard.writeText() 
  }
  function copyToClipboard(text){
    navigator.clipboard.write(text) 
    // document.execCommand
  }
  function selectElementContents(el) {
    var body = document.body, range, sel;
    if (document.createRange && window.getSelection) {
      range = document.createRange();
      sel = window.getSelection();
      sel.removeAllRanges();
      try {
        range.selectNodeContents(el);
        sel.addRange(range);
        // copyToClipboard(range.select());
      } catch (e) {
        range.selectNode(el);
        sel.addRange(range);
        // copyToClipboard(range.select());
      }
    } else if (body.createTextRange) {
      range = body.createTextRange();
      range.moveToElementText(el);
      range.select();
      // copyToClipboard(range.select());
    }
  }
