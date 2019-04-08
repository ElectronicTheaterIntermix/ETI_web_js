var CLIENT_ID = '927628194858-3m94b2s28den641lgsnqna4bik976pek.apps.googleusercontent.com';
var SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly", "https://www.googleapis.com/auth/spreadsheets"];
// var mes = "";
var excel_file="";
// var select = [];
// var btn = [];
var range;
var sentData = [];
var url;

/**
* Check if current user has authorized this application.
*/
function checkAuth() {
    gapi.auth.authorize(
    {
        'client_id': CLIENT_ID,
        'scope': SCOPES.join(' '),
        'immediate': true
    }, handleAuthResult);
}

/**
* Handle response from authorization server.
*
* @param {Object} authResult Authorization result.
*/
function handleAuthResult(authResult) {
    var authorizeDiv = document.getElementById('authorize-div');
    if (authResult && !authResult.error) {
        // Hide auth UI, then load client library.
        authorizeDiv.style.display = 'none';
        loadSheetsApi();
    } else {
        // Show auth UI, allowing the user to initiate authorization by
        // clicking authorize button.
        authorizeDiv.style.display = 'inline';
    }
}

/**
* Initiate auth flow in response to user clicking authorize button.
*
* @param {Event} event Button click event.
*/
function handleAuthClick(event) {
    gapi.auth.authorize(
	{
		client_id: CLIENT_ID, 
		scope: SCOPES, 
		immediate: false
	}, handleAuthResult);
        
    return false;
}

/**
* Load Sheets API client library.
*/
function loadSheetsApi() {
    var discoveryUrl =
    'https://sheets.googleapis.com/$discovery/rest?version=v4';
    // gapi.client.load(discoveryUrl).then(listMajors);
    gapi.client.load(discoveryUrl);
    return discoveryUrl;
    // listMajors();
}

/**
 * Read Sheets and Get Data
 * 
 * @param {callback} it should be "compare" or "setSentData"
 */
function readSheet(callback) {
    var arg1 = arguments[1];
    var arg2 = arguments[2];
    var arg3 = arguments[3];
    gapi.client.sheets.spreadsheets.values.get({
        spreadsheetId: '1hygcWyUd6inFWzYTeVkxMbuU5PLO0bKpEHaBwK4Fn_8',
        range: 'Sheet1!B1:C27',
    }).then(function(response) {
        var range = response.result;
            
        if( callback.name === "compare" ){
            console.log( range.values.length );
            if (range.values.length > 0) {
                callback( excel_file, range.values );
            } else {
                console.log( "No data found");
            }
        }
        else if( callback.name === "setSentData" ){
            callback( arg1, arg2, range.values, arg3 );
        }

    }, function(response) {
        console.log('Error: ' + response.result.error.message);
    });
}

/**
 * Compare the title between excel and google sheet.
 * @param {Array} excel_file : parse excel_file from to_csv()
 * @param {Array} result : data from readSheet() 
 */
function compare( excel_file, result ){
    empty();
    var excel = excel_file.split(","); // ,編劇：紀蔚然；導演：傅裕惠,,
    console.log("Excel : " + excel);
    var index;
    var pre = document.getElementById("ta");

    // 編導
    // excel[4] EX :編劇：紀蔚然；導演：傅裕惠
    if( excel[4] === "" ){
        console.log("It is empty in Excel");
    }else{
        var temp = excel[4].split("；");
        replace(temp,result,0,2);
    }
    
    // 設計
    // excel[5]
    if(excel[5] === ""){
        console.log("It is empty in Excel");  
    }else{
        var temp = excel[5].split("；");
        replace(temp,result,2,13);
    }
    
    // 技術
    // excel[6]
    if(excel[6] === ""){
        console.log("It is empty in Excel");  
    }else{
        var temp = excel[6].split("；");
        replace(temp,result,13,18);
    }

    // 製作/行政
    // excel[8]
    if(excel[8] === ""){
        console.log("It is empty in Excel");  
    }else{
        var temp = excel[8].split("；");
        replace(temp,result,18,19);
    }

    // 表演者
    // excel[7] 
    if(excel[7] === ""){
        console.log("It is empty in Excel");  
    }else{
        var temp = excel[7].split("；");
        replace(temp,result,19,27); 
    }
}

/**
 * Replace the title with those in google sheet.
 * @param {array} temp : excel content split by ";"
 * @param {array} result : data read from google sheet
 * @param {int} start : start from nth row number in google sheet
 * @param {int} end : nth row number which ends in this section
 */
function replace( temp, result, start, end ){
	var index;
	var judge = false;
	var name = ""; // section name ex: 編導, 設計, 技術...
	
	for( var z = 0 ; z < temp.length ; z++ ){
		for( var x = start ; x < end ; x++ ){
			var title = result[x][1].split(",");
			for( var y = 0 ; y < title.length ; y++ ){
				index = temp[z].indexOf( title[y] ); // decide this title should be modified or not
				if( index != -1 ){
					temp[z] = temp[z].replace(title[y],result[x][0]);
					judge = true;
					name = decide_insert( start ); // get the section name 
					insertData( name,temp[z] ); // put content into the box in web page
					break;
				}
			}
		}//x

        if( judge == true ){
            judge = false;
        }
        else{
            console.log( "This is wrong : " + temp[z] + "\n" );
            name = decide_insert( start );
            addSelect( name, temp[z], result ); // add select option to let user classify
        }

	}//z

}

/**
 * Decide the section name.
 * @param {int} num : the start row number to judge which section it belongs.
 */
function decide_insert( num ){
	switch(num){
		case 0:
			return "director";
		case 2:
            return "design";
		case 13:
            return "tech";
		case 18:
            return "actor";
		case 19:
			return "admin";
	}
}

/**
 * Put contents in the webpage's boxes.
 * @param {string} pos : indicates which section should be filled
 * @param {string} data : the contents that will be put in the box
 */
function insertData( pos, data ){
    if(pos === "director"){
        var area = document.works.director;
        area.value = area.value + data + "；";
    }
    if(pos === "design"){
        var area = document.works.design;
        area.value = area.value + data + "；";
    }
    if(pos === "tech"){
        var area = document.works.tech;
        area.value = area.value + data + "；";
    }
    if(pos === "admin"){
        var area = document.works.admin;
        area.value = area.value + data + "；";
    }
    if(pos === "actor"){
        var area = document.works.actor;
        area.value = area.value + data + "；";
    }
}

function empty(){
    var area;
    area = document.works.director;
    area.value = "";
    area = document.works.design;
    area.value = "";
    area = document.works.tech;
    area.value = "";
    area = document.works.admin;
    area.value = "";
    area = document.works.actor;
    area.value = "";
}

/**
 * Add labels and select elements to let users classify the new titles.
 * @param {string} pos : which section it belongs
 * @param {array} data : the content that cannot classify
 * @param {array} result : data read from google sheet
 */
function addSelect( pos, data, result ){
    // add new label to show the content that cannot classify
    var newDiv = document.getElementById('newDiv');
    var titleLabel = document.createElement('label');
    titleLabel.innerHTML = data;
    
    // create select element to let users choose titles
    var titleSelect = document.createElement('select');
    titleSelect.id = data;
    titleLabel.for = data;
    
    // add select options
    switch( pos ){
        case "director":
            for( var x = 0; x < 2; x++ ){
                titleSelect.add( new Option( result[x][0], result[x][0] ) );
            }
            break;
        case "design":
            for( var x = 2; x < 13; x++ ){
                titleSelect.add( new Option( result[x][0], result[x][0] ) );
            }
            break;
        case "tech":
            for( var x = 13; x < 18; x++ ){
                titleSelect.add( new Option( result[x][0], result[x][0] ) ); 
            }
            break;
        case "admin":
            for( var x = 19; x < 27; x++ ){
                titleSelect.add( new Option( result[x][0], result[x][0] ) ); 
            }
            break;
        case "actor":
            for( var x = 18; x < 19; x++ ){
                titleSelect.add( new Option( result[18][0], result[18][0] ) );
            }
            break;
    }
    
    // create button to let user sent data to google sheet
    var btn = document.createElement('button');
    var btnId = data.split("："); // 全形符號
    btn.id = btnId[0]; // the button's id is the title
    btn.type = "button";
    btn.innerHTML = "Enter";

    // it will add new title to google sheet and put it into the box when button is pressed
    btn.addEventListener( "click", function(){
        var url = loadSheetsApi();
        var id = this.id;
        gapi.client.load(url).then( function(){
            readSheet( setSentData, document.getElementById(data), id, pos );
        });
    });
    
    newDiv.appendChild( titleLabel );
    newDiv.appendChild( titleSelect );
    newDiv.appendChild( btn );
    newDiv.appendChild( document.createElement('br') );
}

/**
 * Send data and put data into the webpage's box.
 * @param {object} selectId : selectObject
 * @param {string} btnId : the id of buttons
 * @param {array} result : the data from google sheet
 * @param {string} pos : indicates which section should be filled
 */
function setSentData( selectId, btnId, result, pos ){
    var index = selectId.selectedIndex;
    console.log( "btnId: " + btnId  );
    switch( pos ){
        case "director":
            range = "Sheet1!C" + ( index + 1 ) + ":C" +  ( index + 1 ) ;
            sentData = [[result[index][1] + "," + btnId]];
            console.log(pos);
            break;
        case "design":
            range = "Sheet1!C" + ( index + 3 ) + ":C" + ( index + 3 );
            sentData = [[result[ index + 2 ][1] + "," + btnId]];
            // sentData.push([result[ index + 2 ][1] + "," + btnId]);
            console.log(pos);
            break;
        case "tech":
            range = "Sheet1!C" + ( index + 14 ) + ":C" + ( index + 14 );
            sentData = [[result[ index + 13 ][1] + "," + btnId]];
            console.log(pos);
            break;
        case "admin":
            range = "Sheet1!C" + ( index + 20 ) + ":C" + ( index + 20 );
            sentData = [[result[ index + 19 ][1] + "," + btnId]];
            console.log(pos);
            break;
        case "actor":
            range = "Sheet1!C" + ( index + 19 ) + ":C" + ( index + 19 );
            sentData = [[result[ index + 18 ][1] + "," + btnId]];
            console.log(pos);
            break;
    }

    // replace the title with user's choice
    var text = selectId.id;
    var beReplaced = text.split("：")[0];
    var newTitle = selectId.options[ selectId.selectedIndex ].text;
    text = text.replace( beReplaced, newTitle );

    insertData( pos, text );
    var url =
    'https://sheets.googleapis.com/$discovery/rest?version=v4';
    gapi.client.load(url).then( writeData );
}


/**
 * Write data to sheet.
 * 
 * @param {A1} range 
 * @param {array} message to be written in sheet
 */
function writeData(){
    gapi.client.sheets.spreadsheets.values.update({
       spreadsheetId: '1hygcWyUd6inFWzYTeVkxMbuU5PLO0bKpEHaBwK4Fn_8',
       range: range,
       majorDimension: "ROWS",
       values: sentData,
       valueInputOption: "USER_ENTERED",
    }).then( function(response){
        var range = response.result;
        var rows = range.updatedRows;
        console.log( "updateRows: " + rows );
        var columns = range.updatedColumns;
        console.log( "updateColumns: " + columns );
        var updateRange = range.updatedRange;
        console.log( "updateRange: " + updateRange );
    }, function(response){
        console.log('Error: ' + response.result.error.message );
        
    });
}

/***     Load and parse excel     ***/

function load_JS(filename){
	var fileref;
	fileref = document.createElement('script');
	fileref.setAttribute("type","text/javascript");
	fileref.setAttribute("src",filename);

	if(typeof fileref != "undefined")
		document.getElementsByTagName("head")[0].appendChild(fileref);
}

function createInput(){
	var $input = $('<input id = "fileBtn" type = "file" value = "add file" />');
	$input.appendTo($('body'));

    // add elements which google authorize needs
	///////////////////////////////////////
	var pre = document.createElement('pre');
    pre.id = "output";
    var google = document.createElement('div');
    google.id = "authorize-div";
    google.style.display = 'none';

    var googleBtn = document.createElement('button');
    googleBtn.id = "authorize-button";
	googleBtn.innerHTML = "google assign";
    googleBtn.onClick = handleAuthClick(event);

    google.appendChild(googleBtn);
    document.getElementsByTagName("body")[0].appendChild(google);
    document.getElementsByTagName("body")[0].appendChild(pre);
	//////////////////////////////////////
    
    // create new div to put button and label in addSelect
    var newDiv = document.createElement('div');
    newDiv.id = "newDiv";
    newDiv.style = "position:absolute; top:30; left:900; width:1000"
    document.getElementsByTagName('div')[0].appendChild(newDiv);
}

// parse the contents in excel
function to_csv(workbook){
	var result = [];
	var s_csv = [];
	var sheetName = "ETI資料庫";

	var csv = XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName]);
	s_csv = csv.split("\n");

	var tt = 0;
	var temp = document.getElementsByName("workId");
	var ID = temp[0].getAttribute("value");
	
	for(var x = 0; x < s_csv.length;x++){
		tt = s_csv[x].indexOf(ID);
		if(tt != -1){
			tt = x;
			break;
		}
	}
    excel_file = s_csv[tt];
}

// this function will provide a dialogue box to let user choose its file
function handFile(e){
	var files = e.target.files;
	var num,f;
	for(num = 0, f = files[num]; num != files.length; ++num){
		var reader = new FileReader();
		reader.onload = function(e){
			var data = e.target.result;
			var workbook = XLSX.read(data,{type: "binary"});
			to_csv(workbook);
            url = loadSheetsApi();
            gapi.client.load(url).then( function(){
                readSheet( compare );
            });
		};
		reader.readAsBinaryString(f);
	}
}

// this function is to be sure to load gapi before authorizing and creat fileBtn
function SetKeyCheckAuthority() {
    if(null == gapi.auth) {
        window.setTimeout(SetKeyCheckAuthority,1000);
        return;
    }else{
        checkAuth();
        createInput();
        document.getElementById("fileBtn").addEventListener('change',handFile,false);
    }
}

var curURL = window.location.toString();
console.log("this website is : " + curURL)
if( curURL.indexOf("eti-tw.com/admin/db_format.php?") != -1 ){
	console.log("right website");

	// load libraries which this program needs
	load_JS("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.8.0/jszip.js");
	load_JS("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.8.0/ods.js");
	load_JS("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.8.0/xlsx.core.min.js");
	load_JS("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.8.0/xlsx.js");

    var url = "https://apis.google.com/js/client.js";
    $.getScript( url, function() {
        SetKeyCheckAuthority();
    });
}