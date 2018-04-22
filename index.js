//Excel library
var XLSX = require('xlsx');

//File reading library
var   fs = require('fs');

//Math library
var math = require('mathjs');

//Loging library
var Log = require('log');

//Definition of the logging file
var TimeInMills = new Date().getTime();
var log = new Log('debug', fs.createWriteStream('logs/'+TimeInMills+'.log'));

log.debug("**********************	 	START		****************");

//Location of the Excel File and Workbook
var wb = XLSX.readFileSync("BookSmall.xlsx");
//var wb = XLSX.readFileSync("Book1.xlsx");
var MyWorksheet = wb.Sheets["Sheet1"];

//If Sheet is not found in the workbook
if(MyWorksheet == undefined)
{
	log.debug("No Sheet found");
}
else
{    
	log.debug("Worksheet Exists");
    var InitialRow = 2; //First row to be read from the file
	var	FinalRow = 200; //Last row to be read from the file
    var isEmpty = false;
	
	//Variabled for final loging purposes
	var NumberExpected = FinalRow - InitialRow + 1; //The number of rows that should be properly executed
	var NumberOfClients = 0; //Number of clients created
	var NumberOfCases = 0; //Number of cases created
	var RequestedClients = 0; //Number of clients requested to Bizagi
	var RequestedCases = 0; //Number of cases requested to Bizagi
	
	//Iteratively reading of the file to not make a DoS of the webservices
	CycleMethod(InitialRow, FinalRow, MyWorksheet, function(err, result){});
	
	log.debug('Finished all the rows');
}
log.info("**********************	 	END		****************");
	
/*
*	Itearatively calls itsealf to read the next chunk of the file with a delay to avoid saturating the web services
*	InitialRow: First row that should be read in this cycle
*	FinalRow: Final row that should be read in this cycle
*	MyWorksheet: Excel file to read
*/
function CycleMethod(InitialRow, FinalRow, MyWorksheet, callback) {
	//Last cell of the chunk
	var FinalRowPerCycle = math.min(InitialRow + 50,FinalRow);
	
	//Chunk processing
	ManySOAPCalls(InitialRow, FinalRowPerCycle, MyWorksheet, function(err, result){});
	
	//Updates initial row for next chunk
	InitialRow = FinalRowPerCycle + 1;
	
	//Verify if this should be the last chunk
	if (InitialRow <= FinalRow)
	{	
		//After waiting some time the next chunk will be read
		setTimeout(CycleMethod, 1500, InitialRow, FinalRow, MyWorksheet, function(err, result){});
	}
}
	
function ManySOAPCalls(InitRow, FinRow, worksheet, callback) {
	
	//Definition of the service points
	var soap = require('soap');
	var ServerName = '10.10.10.10';
	var ProjectName = 'PROJECT_NAME';
	var urlWE = 'http://' + ServerName + '/' + ProjectName + '/WebServices/WorkflowEngineSOA.asmx?wsdl';
	var urlEM = 'http://' + ServerName + '/' + ProjectName + '/WebServices/EntityManagerSOA.asmx?wsdl';
	
	//variable to identify of the current chunk should be ended because a empty row was found
	var empty = false;
	log.info('New cycle( ' + InitRow + ' - ' + FinRow + ')');
	console.log('New cycle( ' + InitRow + ' - ' + FinRow + ')');
	
	//Creating of the clients from the service points
	soap.createClientAsync(urlWE).then((clientWE)=>
	{	
	soap.createClientAsync(urlEM).then((clientEM)=>
	{	
		//Cycle to read each of the cells to read
		while(!empty && InitRow <= FinRow) //  && InitRow < 1023
		{		
			//Definition of a variable per each column of the sheet
			ATTRIBUTE_1 = worksheet['A' + InitRow];
			ATTRIBUTE_2 = worksheet['B' + InitRow];
			ATTRIBUTE_3 = worksheet['C' + InitRow];
			ATTRIBUTE_4 = worksheet['I' + InitRow];
			
			//Variable to read if at least one of the cells of the sheet had a variable
			var foundOne = false;
			log.info('Reading Row: '+InitRow);
			
			//For each of the variable we need to verify if at least one has value
			if(ATTRIBUTE_1 != undefined && ATTRIBUTE_1.v != null && ATTRIBUTE_1.v != '') {log.debug('ATTRIBUTE_1 '+ATTRIBUTE_1.v);foundOne = true;}
			if(ATTRIBUTE_2 != undefined && ATTRIBUTE_2.v != null && ATTRIBUTE_2.v != '') {log.debug('ATTRIBUTE_2 '+ATTRIBUTE_2.v);foundOne = true;}
			if(ATTRIBUTE_3 != undefined && ATTRIBUTE_3.v != null && ATTRIBUTE_3.v != '') {log.debug('ATTRIBUTE_3 '+ATTRIBUTE_3.v);foundOne = true;}
			if(ATTRIBUTE_4 != undefined && ATTRIBUTE_4.v != null && ATTRIBUTE_4.v != '') {log.debug('ATTRIBUTE_4 '+ATTRIBUTE_4.v);foundOne = true;}

			//If at least one was found
			if(foundOne)
			{			
				//Definition of the client
				var ClientString = '<M_Clients><sATTRIBUTE_3>$ATTRIBUTE_3</ATTRIBUTE_3><kpTypeofClient><sCode>C</sCode></kpTypeofClient><sATTRIBUTE_2><scode>$ATTRIBUTE_2</scode></sATTRIBUTE_2><ATTRIBUTE_1><sCode>$ATTRIBUTE_1</sCode></ATTRIBUTE_1></M_Clients>';
				
				//Replace the different variables of the xml with the read values
				ClientString = ClientString.replace('$ATTRIBUTE_3', ATTRIBUTE_3.v);
				ClientString = ClientString.replace('$ATTRIBUTE_2', ATTRIBUTE_2.v);
				ClientString = ClientString.replace('$ATTRIBUTE_1', ATTRIBUTE_1.v);
				
				//If the client consents a case will be created at the same time that the client
				if(ATTRIBUTE_4 != undefined && ATTRIBUTE_4.v != null && ATTRIBUTE_4.v != '' && ATTRIBUTE_4.v == 'Y')
				{
					//Formats the Consent and the consent end date
					log.info('Case is going to be created');
					
					//Definition of the string for the case creation
					var CaseCreationString = '<BizAgiWSParam><domain>domain</domain><userName>WebService</userName> <Cases><Case><Process>$PROCESS_NAME</Process><Entities><PROCESS_ENTITY><XIncomingClients>$CLIENTSTRING</XIncomingClients></PROCESS_ENTITY></Entities></Case></Cases></BizAgiWSParam>';
					CaseCreationString = CaseCreationString.replace('$ConsentDate', ConsentYear+'-'+ConsentMonth+'-'+ConsentDay+'T00:00:00.000');
					CaseCreationString = CaseCreationString.replace('$CLIENTSTRING',ClientString);
					
					var args = {casesInfo: CaseCreationString};
					log.debug('Arguments generated');
					log.debug(JSON.stringify(args));

					//Call the create cases web service with the previously generated string
					clientWE.createCasesAsString(args, function(err, result) {
						if(err){
							log.error('Error');
							log.error(err);
							console.log('Error');
							console.log(err);
						}
						if(!result) {

							log.error('Not Working');
							return null;
						}
						log.debug(result);
						//Increase the counters for clients and cases generated
						NumberOfClients ++;
						NumberOfCases ++;
						if(NumberOfClients == NumberExpected)
						{
							log.info('Created Clients: ' + NumberOfClients);
							log.info('Created Cases: ' + NumberOfCases);
							console.log('Created Cases: ' + NumberOfCases);
							console.log('Created Clients: ' + NumberOfClients);
						}
						return(result);
					});
					
					//Increase the requested counters
					RequestedClients++;
					RequestedCases++;
					log.debug('Requested Clients: ' + RequestedClients + 'Requested Cases: ' + RequestedCases );
					
				}
				else
				{
					log.info('Entity is going to be created');
					var args = {entityInfo: '<BizAgiWSParam><Entities>' + ClientString + '</Entities></BizAgiWSParam>'};
					log.debug('Arguments generated');
					log.debug(JSON.stringify(args));

					
					clientEM.saveEntityAsString(args, function(err, result) {
						if(err){
							log.error('Error');
							log.error(err);
							console.log('Error');
							console.log(err);
						}
						if(!result) {

							log.error('Not Working');
							return null;
						}
						log.debug(result);
						//Increase the counters for clients  generated
						NumberOfClients ++;
						if(NumberOfClients == NumberExpected)
						{
							log.info('Created Clients: ' + NumberOfClients);
							log.info('Created Cases: ' + NumberOfCases);
							console.log('Created Cases: ' + NumberOfCases);
							console.log('Created Clients: ' + NumberOfClients);
						}
						return(result);
					});
					//Quest a new client
					RequestedClients++;
					log.debug('Requested Clients: ' + RequestedClients );
				}
				
				InitRow++;
			}
			else
				empty = true;
		}
	});
	});
	log.info('Ended cycle');
	callback(null,!empty);
};