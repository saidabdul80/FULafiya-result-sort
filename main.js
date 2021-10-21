var new_object = []
window.results = [];
var ExcelToJSON = function() {

    this.parseExcel = function(file) {
      var reader = new FileReader();

      reader.onload = function(e) {
        var data = e.target.result;
        var workbook = XLSX.read(data, {
          type: 'binary'
        });
        workbook.SheetNames.forEach(function(sheetName) {          
          var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
          var json_object = JSON.stringify(XL_row_object);
          new_object = [];
          JSON.parse(json_object).map((item, index) => {
              var row = [];
              var y = 0;
              for(let x in item){
                  if(y == 0){
                      y++;
                      row.push(item[x]);
                  }else{
                    row.push(item[x]);
                    if(item[x].includes('/')){
                        row.push(item[x].split('/')[2]);
                    }else{
                        var sp = item[x].split("");
                        var depNamePos;                        
                        for(const i= sp.length; i >= 0 ; i--){
                            if(isNaN(parseInt(sp[i]))){ //is a number, it returns false
                                depNamePos = i;                                
                                break;
                            }
                        }
                        var depName = item[x].slice(depNamePos-3, depNamePos); 
                        row.push(depName);
                    }
                  }
              }
              new_object.push(row);
            });
            var depts = new_object.map(function(item, pos){
              return item[2];
            });
            var uniqueDepts = depts.filter(function(item,pos){
                return  depts.indexOf(item) == pos;
            })
            window.results = [];
            for (let index = 0; index < uniqueDepts.length; index++) {
                window.results[index] = new_object.filter(function(item){
                    if(item[2]== uniqueDepts[index]){
                        return true;
                    }
                });                
            }                      
            document.getElementById("generateK").disabled = false;
         /*  console.log(uniqueDepts);
          console.log(new_object);
          */
          //jQuery( '#xlx_json' ).val( json_object );
        })
      };

      reader.onerror = function(ex) {
        console.log(ex);
      };

      reader.readAsBinaryString(file);
    };
};

function handleFileSelect(evt) {
  
  var files = evt.target.files; // FileList object
  document.getElementById("filenamex").innerHTML = evt.target.value.split('\\')[evt.target.value.split('\\').length-1];
  var xl2json = new ExcelToJSON();
  xl2json.parseExcel(files[0]);
}

function generate(){
    console.log(results);
    
    var wb = XLSX.utils.book_new();
    wb.Props = {
            Title: "Sorted Result By Departments",
            Subject: "Test",
            Author: "Faizat",
            CreatedDate: new Date()
    };
    var ws_data;
    var ws;
    for (let index = 0; index < results.length; index++) {
        wb.SheetNames.push(results[index][0][2]);
        ws_data = results[index];
        ws = XLSX.utils.aoa_to_sheet(ws_data);
        wb.Sheets[results[index][0][2]] = ws;            
    }
    var wbout = XLSX.write(wb, {bookType:'xlsx',  type: 'binary'});
    function s2ab(s) {
            var buf = new ArrayBuffer(s.length);
            var view = new Uint8Array(buf);
            for (var i=0; i<s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
            return buf;
            
    }
    var dx = new Date();
    saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), 'sorted_'+dx+'_.xlsx');
}

(function(){
    document.getElementById('fileToUpload').addEventListener('change', handleFileSelect, false);
    document.getElementById("generateK").addEventListener("click", generate,false);
})()