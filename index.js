//Excel library
var XLSX = require('xlsx');

//File reading library
var   fs = require('fs');

//Math library
var math = require('mathjs');

//Loging library
var Log = require('log');

const utf8 = require('utf8');

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
	var	FinalRow = 2; //Last row to be read from the file
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
	var urlWE = 'http://spa-andress/My_UniCredit/WebServices/WorkflowEngineSOA.asmx?wsdl';
	var urlEM = 'http://spa-andress/My_UniCredit/WebServices/EntityManagerSOA.asmx?wsdl';
	
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
			STATUS = worksheet['A' + InitRow];
			Opu = worksheet['B' + InitRow];
			CNP = worksheet['C' + InitRow];
			CORE_ID = worksheet['D' + InitRow];
			CLIENT_NAME = worksheet['E' + InitRow];
			EMAIL = worksheet['F' + InitRow];
			CLIENT_TYPE = worksheet['G' + InitRow];
			CONSENT_TYPE = worksheet['H' + InitRow];
			CONSENT_VALUE = worksheet['I' + InitRow];
			CONSENT_DATE = worksheet['J' + InitRow];
			CONSENT_END_DATE = worksheet['K' + InitRow];
			CHANNEL = worksheet['L' + InitRow];
			
			//Variable to read if at least one of the cells of the sheet had a variable
			var foundOne = false;
			log.info('Reading Row: '+InitRow);
			
			//For each of the variable we need to verify if at least one has value
			if(STATUS != undefined && STATUS.v != null && STATUS.v != '') {log.debug('STATUS '+STATUS.v);foundOne = true;}
			if(Opu != undefined && Opu.v != null && Opu.v != '') {log.debug('Opu '+Opu.v);foundOne = true;}
			if(CNP != undefined && CNP.v != null && CNP.v != '') {log.debug('CNP '+CNP.v);foundOne = true;}
			if(CORE_ID != undefined && CORE_ID.v != null && CORE_ID.v != '') {log.debug('CORE_ID '+CORE_ID.v);foundOne = true;}
			if(CLIENT_NAME != undefined && CLIENT_NAME.v != null && CLIENT_NAME.v != '') {log.debug('CLIENT_NAME '+CLIENT_NAME.v);foundOne = true;}
			if(EMAIL != undefined && EMAIL.v != null && EMAIL.v != '') {log.debug('EMAIL '+EMAIL.v);foundOne = true;}
			if(CLIENT_TYPE != undefined && CLIENT_TYPE.v != null && CLIENT_TYPE.v != '') {log.debug('CLIENT_TYPE '+CLIENT_TYPE.v);foundOne = true;}
			if(CONSENT_TYPE != undefined && CONSENT_TYPE.v != null && CONSENT_TYPE.v != '') {log.debug('CONSENT_TYPE '+CONSENT_TYPE.v);foundOne = true;}
			if(CONSENT_VALUE != undefined && CONSENT_VALUE.v != null && CONSENT_VALUE.v != '') {log.debug('CONSENT_VALUE '+CONSENT_VALUE.v);foundOne = true;}
			if(CONSENT_DATE != undefined && CONSENT_DATE.v != null && CONSENT_DATE.v != '') {log.debug('CONSENT_DATE '+CONSENT_DATE.v);foundOne = true;}
			if(CONSENT_END_DATE != undefined && CONSENT_END_DATE.v != null && CONSENT_END_DATE.v != '') {log.debug('CONSENT_END_DATE '+CONSENT_END_DATE.v);foundOne = true;}
			if(CHANNEL != undefined && CHANNEL.v != null && CHANNEL.v != '') {log.debug('CHANNEL '+CHANNEL.v);foundOne = true;}

			//If at least one was found
			if(foundOne)
			{			
				//Definition of the client
				var ClientString = '<M_Clients><sCoreId>$CORE_ID</sCoreId><sFullname>$CLIENT_NAME</sFullname><sEmail>$EMAIL</sEmail><sCNP>$CNP</sCNP><kpTypeofClient><sCode>C</sCode></kpTypeofClient><kpBranch><sCode>$Opu</sCode></kpBranch><kpStatus><sCode>$STATUS</sCode></kpStatus></M_Clients>';
				
				//Replace the different variables of the xml with the read values
				ClientString = ClientString.replace('$CORE_ID', CORE_ID.v);
				ClientString = ClientString.replace('$CLIENT_NAME', utf8.encode(CLIENT_NAME.v));
				ClientString = ClientString.replace('$EMAIL', utf8.encode(EMAIL.v));
				ClientString = ClientString.replace('$CNP', CNP.v);
				ClientString = ClientString.replace('$Opu', Opu.v);
				ClientString = ClientString.replace('$STATUS', (STATUS.v == 'ACTIVE')?'AC':(STATUS.v == 'DISABLED')?'CL':'BL');
				
				//If the client consents a case will be created at the same time that the client
				if(CONSENT_VALUE != undefined && CONSENT_VALUE.v != null && CONSENT_VALUE.v != '' && CONSENT_VALUE.v == 'Y')
				{
					//Formats the Consent and the consent end date
					log.info('Case is going to be created');
					var ConsentYear = CONSENT_DATE.v.substring(0,4);
					var ConsentMonth = CONSENT_DATE.v.substring(4,6);
					var ConsentDay = CONSENT_DATE.v.substring(6,8);
					
					var ConsentEndYear = CONSENT_END_DATE.v.substring(0,4);
					var ConsentEndMonth = CONSENT_END_DATE.v.substring(4,6);
					var ConsentEndDay = CONSENT_END_DATE.v.substring(6,8);
					
					//Definition of the string for the case creation
					var CaseCreationString = '<BizAgiWSParam><domain>domain</domain><userName>WebService</userName> <Cases><Case><Process>UpdateClientsConsents</Process><Entities><M_UCP_UpdateConsentReq><dConsentDate>$EndConsentDate</dConsentDate><dConsentEndDate>$ConsentDate</dConsentEndDate><XIncomingClients>$CLIENTSTRING</XIncomingClients><kpIncomingChannel><sCode>C_DB</sCode></kpIncomingChannel>$CONSENT</M_UCP_UpdateConsentReq></Entities></Case></Cases></BizAgiWSParam>';
					CaseCreationString = CaseCreationString.replace('$ConsentDate', ConsentYear+'-'+ConsentMonth+'-'+ConsentDay+'T00:00:00.000');
					CaseCreationString = CaseCreationString.replace('$EndConsentDate', ConsentEndYear+'-'+ConsentEndMonth+'-'+ConsentEndDay+'T00:00:00.000');
					CaseCreationString = CaseCreationString.replace('$CONSENT', '<xConsentsModification><M_ConsentStatus><bNewValue>True</bNewValue><pConsentType><sCode>M</sCode></pConsentType></M_ConsentStatus></xConsentsModification>');
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