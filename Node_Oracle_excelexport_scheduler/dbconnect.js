const oracledb=require('oracledb');
var dbConfig=require('./dbconfig.js');
var excel=require('exceljs');
var scheduler=require('node-schedule');

var j= scheduler.scheduleJob('*/30 * * * *',function(){
async function run(){
	try{
			connection=wait oracledb.getConnection({
			user: dbConfig.user,
			password:dbConfig.password,
			connectString:dbConcifg.connectString
		});

		let result=await connection.execute("select * from fid");
		rows=result.rows;

		if(rows.length!=0){
			var workbook=new excel.Workbook();
			workbook.creator='Me';
			workbook.lastModifiedBy='Her';
			workbook.created=new Date(2019,3,28);
			workbook.modified=new Date();
			workbook.lastPrinted=new Date(2019,3,28);
			workbook.properties.date1904=true;
			workbook.views=[{
				x:0, y:0, width:10000,height:20000,firstSheet:0, activeTab:1, visibility:'visible'
			}]
			var sheet=workbook.addWorksheet('Data');

			for(var i=0;i<rows.length;i++){
				var rowData=[];
				for(var k=0;k<rows[i].length;k++){
					rowData[k+1]=rows[i][k];
				}	
				sheet.addRow(rowData);
			}
			workbook.xlsx.writeFile('I:\\node_module\\excelfiles\\report.xlsx')
				.then(function(){
					console.log("Excel file created successfully");
				});
			}else{
				console.log("No Rows fetched");
		}

	}
	catch(err){
		console.log(err);
	}
	finally{
		if(connection){
			try{
				await connection.close();
				console.log("Connection object closed");
			}
			catch(err){
				console.log(err);
			}
		}
	}

}run()

});