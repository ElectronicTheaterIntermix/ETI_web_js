// Here You can type your custom JavaScript...
var X = XLSX;
var curUrl = window.location.toString();
console.log("url" + curUrl);

function to_csv(workbook) {
	var result = [];
	var s_csv = [];
	var sheetName = "ETI資料庫"
	//workbook.SheetNames.forEach(function(sheetName) {
		var csv = X.utils.sheet_to_csv(workbook.Sheets[sheetName]);
		
		s_csv = csv.split("\n")

        var tt = 0;
		
		for( var x = 0; x < s_csv.length; x++ ){
		    tt = s_csv[x].indexOf("AD001")
		    
		    if( tt != -1 ){
		        tt = x;
		        break;
		        //console.log(tt)
		    }
		}
		
		console.log( s_csv[tt] )
		
		//console.log( typeof(csv) )
		if(csv.length > 0){
			result.push("SHEET: " + sheetName);
			result.push("");
			result.push(csv);
		}
	//});
	return result.join("\n");
}

function createInput(){
    var $input = $('<input id="fileBtn" type="file" value="add file" />');
    $input.appendTo($("body"));
}

function handleFile(e) {
  var files = e.target.files;
  var i,f;
  for (i = 0, f = files[i]; i != files.length; ++i) {
    var reader = new FileReader();
    var name = f.name;
    reader.onload = function(e) {
      var data = e.target.result;
 
      var workbook = XLSX.read(data, {type: 'binary'});
 
      /* DO SOMETHING WITH workbook HERE */
      to_csv(workbook);
    };
    reader.readAsBinaryString(f);
  }
}

if( curUrl.indexOf("http://www.eti-tw.com/admin/db_format.php?") != -1 ){
    createInput();
    document.getElementById("fileBtn").addEventListener('change', handleFile, false);
}


