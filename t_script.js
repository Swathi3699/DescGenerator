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
        if(row["#"]==undefined || row["Identifier"]==undefined || row["From Phase"]==undefined || row["To Phase"]==undefined)
        continue
        else
        {
          num = new String(row['#']).trim().replace('.','')
          htmlData+=`<html><span class="colored" id="demo${i}" onclick=copyToClipboard(document.getElementById("demo${num-1}").textContent)>As per the configured transition #${num} (${row["Identifier"]}), the execution moves from phase '${row["From Phase"]}' to phase '${row["To Phase"]}'. Please refer step 6 Screenshot 'clb_picture_010_Phase Transitions.png'.<br/></span><br/></html>`}
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
    navigator.clipboard.writeText(text)
    
  }