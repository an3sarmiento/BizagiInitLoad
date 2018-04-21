var XLSX = require('xlsx');
var   fs = require('fs');
var math = require('mathjs');
var Log = require('log');
var TimeInMills = new Date().getTime();
var log = new Log('debug', fs.createWriteStream('logs/'+TimeInMills+'.log'));
log.debug("**********************	 	START		****************");


var wb = XLSX.readFileSync("Book1 - copia.xlsx");
//var wb = XLSX.readFileSync("Book1.xlsx");

var MyWorksheet = wb.Sheets["Sheet2"];
log.debug("In");
if(MyWorksheet == undefined)
{
	
}
else
{    
	log.debug("Worksheet Exists");
    var InitialRow = 2
	var	FinalRow = 140000;
    var isEmpty = false;

	var NumberOfExecuted = 0;
	var NumberOfWaiting = 0;
	var cycles = 0;
	var NumberExpected = FinalRow - InitialRow + 1;
	var NumberOfClients = 0;
	var NumberOfCases = 0;
	var RequestedClients = 0;
	var RequestedCases = 0;
	
	CycleMethod(InitialRow, FinalRow, MyWorksheet, cycles, function(err, result){});
	
	log.debug('Finished all the rows');
	log.info("**********************	 	END		****************");
}
	
function CycleMethod(InitialRow, FinalRow, MyWorksheet, cycles, callback) {
	var FinalRowPerCycle = math.min(InitialRow + 50,FinalRow);
	
	//setTimeout(ManySOAPCalls, 0, InitialRow, FinalRowPerCycle, MyWorksheet, cycles, function(err, result){})
	ManySOAPCalls(InitialRow, FinalRowPerCycle, MyWorksheet, cycles, function(err, result){});
	InitialRow = FinalRowPerCycle + 1;

	cycles++;
	//isEmpty= true;
	if (InitialRow <= FinalRow)
	{	
		setTimeout(CycleMethod, 1500, InitialRow, FinalRow, MyWorksheet, cycles, function(err, result){});
	}
}
	
function ManySOAPCalls(InitRow, FinRow, worksheet, numberOfSeconds, callback) {
	

	var soap = require('soap');
	var urlWE = 'http://spa-andress/My_UniCredit/WebServices/WorkflowEngineSOA.asmx?wsdl';
	var urlEM = 'http://spa-andress/My_UniCredit/WebServices/EntityManagerSOA.asmx?wsdl';
	empty = false;
	log.info('New cycle( ' + InitRow + ' - ' + FinRow + ')');
	console.log('New cycle( ' + InitRow + ' - ' + FinRow + ')');
	
	soap.createClientAsync(urlWE).then((clientWE)=>
	{	
	soap.createClientAsync(urlEM).then((clientEM)=>
	{	
		while(!empty && InitRow <= FinRow) //  && InitRow < 1023
		{		
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
			
			var foundOne = false;
			log.info('Reading Row: '+InitRow);
			
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

			
			if(foundOne)
			{			
				InitRow++;
				NumberOfWaiting++;
				
				var ClientString = '<M_Clients><sCoreId>$CORE_ID</sCoreId><sFullname>$CLIENT_NAME</sFullname><sEmail>$EMAIL</sEmail><sCNP>$CNP</sCNP><kpTypeofClient><sCode>C</sCode></kpTypeofClient><kpBranch><sCode>$Opu</sCode></kpBranch><kpStatus><sCode>$STATUS</sCode></kpStatus></M_Clients>';
				
				ClientString = ClientString.replace('$CORE_ID', CORE_ID.v);
				ClientString = ClientString.replace('$CLIENT_NAME', CLIENT_NAME.v);
				ClientString = ClientString.replace('$EMAIL', EMAIL.v);
				ClientString = ClientString.replace('$CNP', CNP.v);
				ClientString = ClientString.replace('$Opu', Opu.v);
				ClientString = ClientString.replace('$STATUS', (STATUS.v == 'ACTIVE')?'AC':(STATUS.v == 'DISABLED')?'CL':'BL');
				
				if(CONSENT_VALUE != undefined && CONSENT_VALUE.v != null && CONSENT_VALUE.v != '' && CONSENT_VALUE.v == 'Y')
				{
					log.info('Case is going to be created');
					var ConsentYear = CONSENT_DATE.v.substring(0,4);
					var ConsentMonth = CONSENT_DATE.v.substring(4,6);
					var ConsentDay = CONSENT_DATE.v.substring(6,8);
					
					var ConsentEndYear = CONSENT_END_DATE.v.substring(0,4);
					var ConsentEndMonth = CONSENT_END_DATE.v.substring(4,6);
					var ConsentEndDay = CONSENT_END_DATE.v.substring(6,8);
					
					var CaseCreationString = '<BizAgiWSParam><domain>domain</domain><userName>WebService</userName> <Cases><Case><Process>UpdateClientsConsents</Process><Entities><M_UCP_UpdateConsentReq><dConsentDate>$EndConsentDate</dConsentDate><dConsentEndDate>$ConsentDate</dConsentEndDate><XIncomingClients>$CLIENTSTRING</XIncomingClients><kpIncomingChannel><sCode>C_DB</sCode></kpIncomingChannel>$CONSENT</M_UCP_UpdateConsentReq></Entities></Case></Cases></BizAgiWSParam>';
					CaseCreationString = CaseCreationString.replace('$ConsentDate', ConsentYear+'-'+ConsentMonth+'-'+ConsentDay+'T00:00:00.000');
					CaseCreationString = CaseCreationString.replace('$EndConsentDate', ConsentEndYear+'-'+ConsentEndMonth+'-'+ConsentEndDay+'T00:00:00.000');
					CaseCreationString = CaseCreationString.replace('$CONSENT', '<xConsentsModification><M_ConsentStatus><bNewValue>True</bNewValue><pConsentType><sCode>M</sCode></pConsentType></M_ConsentStatus></xConsentsModification>');
					CaseCreationString = CaseCreationString.replace('$CLIENTSTRING',ClientString);
					
					var args = {casesInfo: CaseCreationString};
					log.debug('Arguments generated');
					log.debug(JSON.stringify(args));

					
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
					RequestedClients++;
					log.debug('Requested Clients: ' + RequestedClients );
				}
			}
			else
				empty = true;
		}
	});
	});
	log.info('Ended cycle');
	callback(null,!empty);
};