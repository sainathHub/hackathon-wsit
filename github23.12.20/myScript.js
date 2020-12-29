'use strict';
	var filesdata;

	


	;( function ( document, window, index )
	{
		// feature detection for drag&drop upload
		var isAdvancedUpload = function()
			{
				var div = document.createElement( 'div' );
				return ( ( 'draggable' in div ) || ( 'ondragstart' in div && 'ondrop' in div ) ) && 'FormData' in window && 'FileReader' in window;
			}();


		// applying the effect for every form
		var forms = document.querySelectorAll( '.box' );
		Array.prototype.forEach.call( forms, function( form )
		{
			var input		 = form.querySelector( 'input[type="file"]' ),
				label		 = form.querySelector( 'label' ),
				errorMsg	 = form.querySelector( '.box__error span' ),
				restart		 = form.querySelectorAll( '.box__restart' ),
				droppedFiles = false,
				showFiles	 = function( files )
				{
					label.textContent = files.length > 1 ? ( input.getAttribute( 'data-multiple-caption' ) || '' ).replace( '{count}', files.length ) : files[ 0 ].name;
				},
				triggerFormSubmit = function()
				{
					var event = document.createEvent( 'HTMLEvents' );
					event.initEvent( 'submit', true, false );
					form.dispatchEvent( event );
				};



			// automatically submit the form on file select
			input.addEventListener( 'change', function( e )
			{
				showFiles( e.target.files );
				console.log(e.target.files);
				Upload(e.target.files, false)
				triggerFormSubmit()
			});

			// drag&drop files if the feature is available
			if( isAdvancedUpload )
			{
				form.classList.add( 'has-advanced-upload' ); // letting the CSS part to know drag&drop is supported by the browser

				[ 'drag', 'dragstart', 'dragend', 'dragover', 'dragenter', 'dragleave', 'drop' ].forEach( function( event )
				{
					form.addEventListener( event, function( e )
					{
						// preventing the unwanted behaviours
						e.preventDefault();
						e.stopPropagation();
					});
				});
				[ 'dragover', 'dragenter' ].forEach( function( event )
				{
					form.addEventListener( event, function()
					{
						form.classList.add( 'is-dragover' );
					});
				});
				[ 'dragleave', 'dragend', 'drop' ].forEach( function( event )
				{
					form.addEventListener( event, function()
					{
						form.classList.remove( 'is-dragover' );
					});
				});
				form.addEventListener( 'drop', function( e )
				{
				
					droppedFiles = e.dataTransfer.files; // the files that were dropped
					filesdata = droppedFiles
					showFiles( droppedFiles );
					console.log(droppedFiles[0].name);
					Upload(false);

					triggerFormSubmit();

									});
			}


			// if the form was submitted
			form.addEventListener( 'submit', function( e )
			{
				// preventing the duplicate submissions if the current one is in progress
				if( form.classList.contains( 'is-uploading' ) ) return false;

				form.classList.add( 'is-uploading' );
				form.classList.remove( 'is-error' );

			});


			// restart the form if has a state of error/success	
			Array.prototype.forEach.call( restart, function( entry )
			{
				entry.addEventListener( 'click', function( e )
				{
					e.preventDefault();
					form.classList.remove( 'is-error', 'is-success' );
					input.click();
				});
			});
			// Firefox focus bug fix for file input
			input.addEventListener( 'focus', function(){ input.classList.add( 'has-focus' ); });
			input.addEventListener( 'blur', function(){ input.classList.remove( 'has-focus' ); });

		});
	}( document, window, 0 ));


var to_html = (workbook) => {
	var HTMLOUT = document.getElementById("dvExcel")
	HTMLOUT.visibility = "hidden";
	HTMLOUT.innerHTML = "";
	workbook.SheetNames.forEach(function (sheetName) {
		var htmlstr = XLSX.write(workbook, { sheet: sheetName, type: 'string', bookType: 'html' });
		HTMLOUT.innerHTML += htmlstr;
	});
	var table = document.getElementsByClassName('div.table')
	table.className = 'tablelizer-table';
	console.log(HTMLOUT)
	return "HTMLOUT";
};
function Upload(flag) {
		var fileName = filesdata[0].name
		//Reference the FileUpload element.
		//Validate whether File is valid Excel file.
		var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xls|.xlsx)$/;
		if (regex.test(fileName.toLowerCase())) {
			if (typeof (FileReader) != "undefined") {
				var reader = new FileReader();
 
				//For Browsers other than IE.
				if (reader.readAsBinaryString) {
					reader.onload = function (e) {
						console.log(e.target.result);
						if(flag == false)
						processExcel(e.target.result);
						else
						ParseExcel(e.target.result);
					};
					reader.readAsBinaryString(filesdata[0]);
				} else {
					//For IE Browser.
					reader.onload = function (e) {
						var data = "";
						var bytes = new Uint8Array(e.target.result);
						for (var i = 0; i < bytes.byteLength; i++) {
							data += String.fromCharCode(bytes[i]);
						}
						if(flag == false)
						processExcel(data);
						else
						ParseExcel(data);
					};
					reader.readAsArrayBuffer(filesdata[0]);
				}
			} else {
				alert("This browser does not support HTML5.");
			}
		} else {
			alert("Please upload a valid Excel file.");
		}
	};
function createTable(flag) {
	var table = document.createElement("table");
		table.border = "1";
		table.className = "tableizer-table"
		//Add the header row.
		var row = table.insertRow(-1);
 
		//Add the header cells.
		if(flag==0){
		var headerCell = document.createElement("TH");
		headerCell.innerHTML = "plus 1";
		row.appendChild(headerCell);
		}
		headerCell = document.createElement("TH");
		headerCell.innerHTML = "Recepient Type";
		row.appendChild(headerCell);
 
		headerCell = document.createElement("TH");
		headerCell.innerHTML = "Reason for GAndE";
		row.appendChild(headerCell);
	if(flag==0){
		headerCell = document.createElement("TH");
		headerCell.innerHTML = "Recepient details";
		row.appendChild(headerCell);
	}
	else{
		headerCell = document.createElement("TH");
		headerCell.innerHTML = "Forename";
		row.appendChild(headerCell);
		headerCell = document.createElement("TH");
		headerCell.innerHTML = "Surname";
		row.appendChild(headerCell);
	}

		headerCell = document.createElement("TH");
		headerCell.innerHTML = "Organization";
		row.appendChild(headerCell);

		 
		headerCell = document.createElement("TH");
		headerCell.innerHTML = "Role In Organisation";
		row.appendChild(headerCell);
	if (flag==0){
		headerCell = document.createElement("TH");
		headerCell.innerHTML = "Associated with plus 1";
		row.appendChild(headerCell);

		headerCell = document.createElement("TH");
		headerCell.innerHTML = "Agregated value per Recipient";
		row.appendChild(headerCell);

		
		headerCell = document.createElement("TH");
		headerCell.innerHTML = "index";
		row.appendChild(headerCell);
	}
		return table;
	}
var index = 1;
	function add_nonplus(table, excelRows,  i){
		    var isplus = 0;
			var pluscount = 0;
			for (var j=i+1; j<excelRows.length && excelRows[j].Type.includes("Plus 1"); ++j, pluscount++ );
			var row = table.insertRow(-1);
				//Add the data cells.
			if (pluscount > 0){
				isplus = 1;
			}
				
			var cell = row.insertCell(-1);
			cell.innerHTML = "plus 1"+ "\(" + pluscount + "\)";
			cell = row.insertCell(-1);
			cell.innerHTML = excelRows[i].Type;
			cell = row.insertCell(-1);
			cell.innerHTML = excelRows[i].ReasonForGAndE;
			cell = row.insertCell(-1);
			cell.innerHTML = excelRows[i].Forename.concat(" ".concat(excelRows[i].Surname));
			cell = row.insertCell(-1);
			cell.innerHTML = excelRows[i].Organization;
			cell = row.insertCell(-1);
			cell.innerHTML = excelRows[i].Role;
			cell = row.insertCell(-1);
			cell.innerHTML = isplus;
			cell = row.insertCell(-1);
			cell.innerHTML = "0.0$";
			cell = row.insertCell(-1);
			cell.innerHTML = index;

			index++;

			return pluscount;
	};

	function add_plus1(table, excelRows, i) {
		
		var row = table.insertRow(-1);
			var cell = row.insertCell(-1);
			cell.style.visibility = "hidden";
 
			cell = row.insertCell(-1);
			cell.innerHTML = excelRows[i].Forename;
			
			var cell = row.insertCell(-1);
			cell.innerHTML = "";
			
			cell = row.insertCell(-1);
			cell.innerHTML = excelRows[i].Surname;
			
			row.insertCell(-1);
			row.insertCell(-1);
			row.insertCell(-1);
			row.insertCell(-1);

	};


	function add_header(table){
		var row = table.insertRow(-1);
			var cell1 = row.insertCell(0);
			cell1.style.visibility = "hidden";
			var cell2 = row.insertCell(1);
			cell2.innerHTML = "Forename";
			
			var cell1 = row.insertCell(2);
			
			var cell1 = row.insertCell(3);
			cell1.innerHTML = "surname";

			row.insertCell(-1);
			row.insertCell(-1);
			row.insertCell(-1);
			row.insertCell(-1);
	}


	function ParseExcel(data) {
		//Read the Excel File data.
		
		var workbook = XLSX.read(data, {
			type: 'binary'
		});
 
		//Fetch the name of First Sheet.
		var firstSheet = workbook.SheetNames[0];
		
		//Read all rows from First Sheet into an JSON array.
		var excelRows = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[firstSheet]);
		console.log(excelRows)
		//Create a HTML Table element.
		var table = createTable(0);
		console.log("created table");
		//Add the data rows from Excel file.

		console.log(table);
		var dvparsed = document.getElementById("dvparsed");
		dvparsed.innerHTML = "";
		dvparsed.appendChild(table);
	};

	function processExcel(data) {
				//Read the Excel File data.
		var workbook = XLSX.read(data, {
			type: 'binary'
		});
 
		//Fetch the name of First Sheet.
		var firstSheet = workbook.SheetNames[0];
 
		//Read all rows from First Sheet into an JSON array.
		var excelRows = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[firstSheet]);

		//Create a HTML Table element.
		var table = createTable(1);
		console.log("created table");
		//Add the data rows from Excel file.
		for (var i = 0; i < excelRows.length && typeof excelRows[i].Type != "undefined"; i++) {
	
			
			if (excelRows[i].Type.includes("Plus 1")) {
				
			var row = table.insertRow(-1);
			var cell = row.insertCell(-1);
			cell.innerHTML= excelRows[i].Type;
 
			var cell = row.insertCell(-1);
			cell.innerHTML = "";

			cell = row.insertCell(-1);
			cell.innerHTML = excelRows[i].Forename;
			cell = row.insertCell(-1);
			cell.innerHTML = excelRows[i].Surname;
			
			row.insertCell(-1);
			row.insertCell(-1);

			console.log("created nonplus");

			}
  
			else {
				var row = table.insertRow(-1);	
				var cell = row.insertCell(-1);
				cell.innerHTML = excelRows[i].Type;
				cell = row.insertCell(-1);
				cell.innerHTML = excelRows[i].ReasonForGAndE;
				cell = row.insertCell(-1);
				cell.innerHTML = excelRows[i].Forename;
				cell = row.insertCell(-1);
				cell.innerHTML = excelRows[i].Surname;
				cell = row.insertCell(-1);
				cell.innerHTML = excelRows[i].Organization;
				cell = row.insertCell(-1);
				cell.innerHTML = excelRows[i].Role;

			}
		}
		console.log(table);
		var c = document.getElementById('form-main').children;
		var i;
		for (i = 0; i < c.length; i++) {
		c[i].style.display = "none";
		}
		var dvExcel = c[c.length-1];
		dvExcel.style.display = "inline-block"
		dvExcel.innerHTML = "";
		dvExcel.appendChild(table);
	};