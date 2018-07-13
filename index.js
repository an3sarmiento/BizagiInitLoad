//Excel library
var XLSX = require('xlsx');

//File reading library
var   fs = require('fs');

//Math library
var math = require('mathjs');

//Loging library
var LogF = require('log');

//Definition of the service points
var SoapServer = require('soap');

//Cycle related variables
var jumpSize, basicTime,cyclesSinceLastStop, files, jumpsFile, ErrorJson = [], WaitingTime;

//Variabled for final loging purposes
var NumberExpected, NumberOfClients, NumberOfCases, NumberOfErrorClients, NumberOfErrorCases, RequestedClients, RequestedCases, AppError, TimeInMills;
	
var RemainingClients = 0;
var RemainingCases = 0;
var maxClients = 400;
var minTime = 10;
var minTH = 180;
var maxTH = 220;
		

//var SERVER_NAME = 'bizagidev';
//var SERVER_NAME = 'digitalflowtest';
var SERVER_NAME = 'bizagidev';
var APPLICATION_NAME = 'My_UniCredit';
var urlWE = 'http://' + SERVER_NAME + '/' + APPLICATION_NAME + '/WebServices/WorkflowEngineSOA.asmx?wsdl';
var urlEM = 'http://' + SERVER_NAME + '/' + APPLICATION_NAME + '/WebServices/EntityManagerSOA.asmx?wsdl';

var ExcelsToRead = {excels:[
	{name:"RECONCILIATION FILE1.xlsx",sheet:"SHEET1",init:0,fin:72769}
]};
var ExcelPos = 0;
NextExcel();

function NextExcel()
{
	if(ExcelsToRead.excels.length > ExcelPos)
	{	
		var MyExcel = ExcelsToRead.excels[ExcelPos]
		ReadExcel(MyExcel.name,MyExcel.sheet, MyExcel.init, MyExcel.fin,ExecuteError);
	}
	ExcelPos++;
}


function ReadExcel(FileName, SheetName, HeaderRow, FinalRow, callback)
{
	
	//Definition of the logging file
	TimeInMills = new Date().getTime();
	
	Logs = {
		logFDebug : new LogF('debug', fs.createWriteStream('logs/'+SERVER_NAME+'/'+TimeInMills+'-debug.log')),
		logFInfo : new LogF('info', fs.createWriteStream('logs/'+SERVER_NAME+'/'+TimeInMills+'-info.log')),
		logFError : new LogF('error', fs.createWriteStream('logs/'+SERVER_NAME+'/'+TimeInMills+'-error.log'))
	};

	jumpSize = 5;
	basicTime = 100;
	cyclesSinceLastStop = 0;
	files = 1;
	jumpsFile = 0;
	AppError = false;
	
	ErrorJson = [];
	Log(Logs,'info',"***********	 	START	" + FileName  + "	**********");
	
	//Location of the Excel File and Workbook
	var wb = XLSX.readFileSync(FileName);
	var MyWorksheet = wb.Sheets[SheetName];
		
	//If Sheet is not found in the workbook
	if(MyWorksheet == undefined)
	{
		Log(Logs,'info',"No Sheet found");
	}
	else
	{    
		Log(Logs,'debug',"Worksheet Exists");
		
		//Variabled for final loging purposes
		NumberExpected = FinalRow - HeaderRow; //The number of rows that should be properly executed
		NumberOfClients = 0; //Number of clients created
		NumberOfCases = 0; //Number of cases created
		NumberOfErrorClients = 0; //Number of clients created
		NumberOfErrorCases = 0; //Number of cases created
		RequestedClients = 0; //Number of clients requested to Bizagi
		RequestedCases = 0; //Number of cases requested to Bizagi
		RemainingClients = 0;
		RemainingCases = 0;

		//Iteratively reading of the file to not make a DoS of the webservices
		CycleMethod(Logs, HeaderRow + 1, FinalRow, MyWorksheet, callback);
	}
}
/*
*	Itearatively calls itsealf to read the next chunk of the file with a delay to avoid saturating the web services
*	InitialRow: First row that should be read in this cycle
*	FinalRow: Final row that should be read in this cycle
*	MyWorksheet: Excel file to read
*/
//function CycleMethod(InitialRow, FinalRow, MyWorksheet, clientEM, clientWE, callback) {
function CycleMethod(LogsFiles, InitialRow, FinalRow, MyWorksheet, callback) {
	
	var NewRemainingClients = RequestedClients - NumberOfClients - NumberOfErrorClients;
	var NewRemainingCases = RequestedCases - NumberOfCases - NumberOfErrorCases;
	
	var Grew = math.max(math.min(NewRemainingClients - RemainingClients,10),-10)/5;	
	
	RemainingClients = NewRemainingClients;
	RemainingCases = NewRemainingCases;
	
	//Last cell of the chunk
	var FinalRowPerCycle = math.min(InitialRow + jumpSize - 1,FinalRow);
	
	if(jumpsFile > 1000)
	{
		LogsFiles.logFDebug = new LogF('debug', fs.createWriteStream('logs/'+SERVER_NAME+'/'+TimeInMills+'-debug (' + files + ').log'));
		jumpsFile = 0;
		files++;
	}
	jumpsFile ++
	
	//Chunk processing
	ManySOAPCalls(LogsFiles, InitialRow, FinalRowPerCycle, MyWorksheet, callback);
	//ManySOAPCalls(InitialRow, FinalRowPerCycle, MyWorksheet, clientEM, clientWE, callback);
	
	//Updates initial row for next chunk
	InitialRow = FinalRowPerCycle;
	
	WaitingTime = basicTime;
	cyclesSinceLastStop ++;
	if(AppError)
	{
		AppError = false;
		WaitingTime = 60000;
		Log(LogsFiles, 'error',"\nStop app error");
		cyclesSinceLastStop = 0;
	}
	else if(RemainingCases > 400 || RemainingClients > maxClients)
	{
		basicTime += 100;
		if(cyclesSinceLastStop < 50)
			basicTime += 100;
		Log(LogsFiles, 'info',"\nStop for big number of cases");
		Log(LogsFiles, 'info',"Cycles Since Last Stop: " + cyclesSinceLastStop + "\n");
		WaitingTime = 20000;
		cyclesSinceLastStop = 0;
	}
	else if(Grew <= 0 && RemainingClients < minTH && cyclesSinceLastStop > 20)
	{
		basicTime -= (minTH/(minTH-RemainingClients))^(1+Grew);
	}
	else if(Grew >= 0 && RemainingClients > maxTH)
	{
		basicTime += (((maxClients-maxTH)/(maxClients-RemainingClients))^(1+Grew));
	}
	basicTime = math.max(minTime,parseInt(basicTime));
	
	//Verify if this should be the last chunk
	if (InitialRow < FinalRow)
	{	
		//After waiting some time the next chunk will be read
		//setTimeout(CycleMethod, WaitingTime, InitialRow + 1 , FinalRow, MyWorksheet, clientEM, clientWE, function(err, result){});
		setTimeout(CycleMethod, WaitingTime, LogsFiles, InitialRow + 1 , FinalRow, MyWorksheet, function(err, result){});
	}
	else
	{
		setTimeout(NextExcel, 10000);
	}
}
	
//function ManySOAPCalls(InitRow, FinRow, worksheet, clientEM, clientWE, callback) {
function ManySOAPCalls(LogsFiles, InitRow, FinRow, worksheet, callback) {
	
		
	SoapServer.createClientAsync(urlEM).then((clientEM)=>
	{
		SoapServer.createClientAsync(urlWE).then((clientWE)=>
		{
				
		//variable to identify of the current chunk should be ended because a empty row was found
		var empty = false;
		Log(LogsFiles, 'info','New cycle( ' + InitRow + ' - ' + FinRow + ')\t\tWaiting clients: ' + (RequestedClients - NumberOfClients - NumberOfErrorClients) + ', basicTime ' + basicTime );
				
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
			Log(LogsFiles, 'debug','Reading Row: '+InitRow);
			
			//For each of the variable we need to verify if at least one has value
			if(STATUS != undefined && STATUS.v != null && STATUS.v != '') {Log(LogsFiles, 'debug','STATUS '+STATUS.v);foundOne = true;}
			if(Opu != undefined && Opu.v != null && Opu.v != '') {Log(LogsFiles, 'debug','Opu '+Opu.v);foundOne = true;}
			if(CNP != undefined && CNP.v != null && CNP.v != '') {Log(LogsFiles, 'debug','CNP '+CNP.v);foundOne = true;}
			if(CORE_ID != undefined && CORE_ID.v != null && CORE_ID.v != '') {Log(LogsFiles, 'debug','CORE_ID '+CORE_ID.v);foundOne = true;}
			if(CLIENT_NAME != undefined && CLIENT_NAME.v != null && CLIENT_NAME.v != '') {Log(LogsFiles, 'debug','CLIENT_NAME '+CLIENT_NAME.v);foundOne = true;}
			if(EMAIL != undefined && EMAIL.v != null && EMAIL.v != '') {Log(LogsFiles, 'debug','EMAIL '+EMAIL.v);foundOne = true;}
			if(CLIENT_TYPE != undefined && CLIENT_TYPE.v != null && CLIENT_TYPE.v != '') {Log(LogsFiles, 'debug','CLIENT_TYPE '+CLIENT_TYPE.v);foundOne = true;}
			if(CONSENT_TYPE != undefined && CONSENT_TYPE.v != null && CONSENT_TYPE.v != '') {Log(LogsFiles, 'debug','CONSENT_TYPE '+CONSENT_TYPE.v);foundOne = true;}
			if(CONSENT_VALUE != undefined && CONSENT_VALUE.v != null && CONSENT_VALUE.v != '') {Log(LogsFiles, 'debug','CONSENT_VALUE '+CONSENT_VALUE.v);foundOne = true;}
			if(CONSENT_DATE != undefined && CONSENT_DATE.v != null && CONSENT_DATE.v != '') {Log(LogsFiles, 'debug','CONSENT_DATE '+CONSENT_DATE.v);foundOne = true;}
			if(CONSENT_END_DATE != undefined && CONSENT_END_DATE.v != null && CONSENT_END_DATE.v != '') {Log(LogsFiles, 'debug','CONSENT_END_DATE '+CONSENT_END_DATE.v);foundOne = true;}
			if(CHANNEL != undefined && CHANNEL.v != null && CHANNEL.v != '') {Log(LogsFiles, 'debug','CHANNEL '+CHANNEL.v);foundOne = true;}

			//If at least one was found
			if(foundOne)
			{			
				//Definition of the client
				var ClientString = '<sCoreId>$CORE_ID</sCoreId><sFullname>$CLIENT_NAME</sFullname><sEmail>$EMAIL</sEmail><sCNP>$CNP</sCNP>$Type$Opu<kpStatus><sCode>$STATUS</sCode></kpStatus>';
				
				//Replace the different variables of the xml with the read values
				ClientString = ClientString.replace('$CORE_ID', CORE_ID.v);
				var CN = (CLIENT_NAME != undefined && CLIENT_NAME.v != null && CLIENT_NAME.v != '')?CLIENT_NAME.v:"";
				var CE = (EMAIL != undefined && EMAIL.v != null && EMAIL.v != '')?EMAIL.v:"";
				var BR = (Opu != undefined && Opu.v != null && Opu.v != '')?'<kpBranch><sCode>' + Opu.v + '</sCode></kpBranch>':'';
				ClientString = ClientString.replace('$CLIENT_NAME', CleanText(CN));
				//ClientString = ClientString.replace('$EMAIL', EMAIL.v);
				ClientString = ClientString.replace('$EMAIL', CleanText(CE));
				ClientString = ClientString.replace('$CNP', CNP.v);
				ClientString = ClientString.replace('$Opu', BR);
				//ClientString = ClientString.replace('$Type', '<kpTypeofClient><sCode>C</sCode></kpTypeofClient>');
				ClientString = ClientString.replace('$Type', '');
				ClientString = ClientString.replace('$STATUS', (STATUS == undefined || STATUS.v == null || STATUS.v == '' || STATUS.v == 'ACTIVE')?'AC':(STATUS.v == 'B')?'BL':'CL');
				
				var XMLCoding = '<?xml version="1.0" encoding="ISO-8859-1"?>';
				
				//If the client consents a case will be created at the same time that the client
				if(CONSENT_VALUE != undefined && CONSENT_VALUE.v != null && CONSENT_VALUE.v != '' && (CONSENT_VALUE.v == 'Y' || CONSENT_VALUE.v == 'N'))
				{
					
					//Formats the Consent and the consent end date
					Log(LogsFiles, 'debug','Case is going to be created');
					var CONSENT_DATE_STRING = '' + CONSENT_DATE.v;
					var ConsentYear = CONSENT_DATE_STRING.substring(0,4);
					var ConsentMonth = CONSENT_DATE_STRING.substring(4,6);
					var ConsentDay = CONSENT_DATE_STRING.substring(6,8);
					
					var CONSENT_DATE_STRING = '' + CONSENT_END_DATE.v;
					var ConsentEndYear = CONSENT_DATE_STRING.substring(0,4);
					var ConsentEndMonth = CONSENT_DATE_STRING.substring(4,6);
					var ConsentEndDay = CONSENT_DATE_STRING.substring(6,8);
					
					//Definition of the string for the case creation
					var CaseCreationString = '<![CDATA[$XMLCoding<BizAgiWSParam><domain>domain</domain><userName>WebService</userName> <Cases><Case><Process>UpdateClientsConsents</Process><Entities><M_UCP_UpdateConsentReq><dConsentDate>$ConsentDate</dConsentDate><dConsentEndDate>$EndConsentDate</dConsentEndDate><kmClient>$CLIENTSTRING</kmClient><kpIncomingChannel><sCode>C_DB</sCode></kpIncomingChannel><xConsentsModification><M_ConsentStatus><bNewValue>$CONSENT_VALUE</bNewValue><pConsentType><sCode>M</sCode></pConsentType></M_ConsentStatus></xConsentsModification></M_UCP_UpdateConsentReq></Entities></Case></Cases></BizAgiWSParam>]]>';
					CaseCreationString = CaseCreationString.replace('$ConsentDate', ConsentYear+'-'+ConsentMonth+'-'+ConsentDay+'T00:00:00.000');
					CaseCreationString = CaseCreationString.replace('$EndConsentDate', ConsentEndYear+'-'+ConsentEndMonth+'-'+ConsentEndDay+'T00:00:00.000');
					CaseCreationString = CaseCreationString.replace('$CONSENT_VALUE', (CONSENT_VALUE.v=='Y')?'True':'False');
					CaseCreationString = CaseCreationString.replace('$CLIENTSTRING',ClientString);
					CaseCreationString = CaseCreationString.replace('$XMLCoding',XMLCoding);
					
					var args = {casesInfo: CaseCreationString};

					var properties = {
							'ClientArguments':args,
							'isCase': true,
							'RowNumber': InitRow,
							'retry': 0,
							'client': clientWE,
							'callback': callback,
							'LogsFiles' : LogsFiles
					}
					//Call the create cases web service with the previously generated string				
					OneSOAPCall(properties);
					
					//Increase the requested counters
					RequestedClients++;
					RequestedCases++;
					Log(LogsFiles, 'debug','Requested Clients: ' + RequestedClients + '\tRequested Cases: ' + RequestedCases );
					
					
				}
				else
				{
					//Formats the Consent and the consent end date
					Log(LogsFiles, 'debug','Case is going to be created');
					var CONSENT_DATE_STRING = '' + CONSENT_DATE.v;
					var ConsentYear = CONSENT_DATE_STRING.substring(0,4);
					var ConsentMonth = CONSENT_DATE_STRING.substring(4,6);
					var ConsentDay = CONSENT_DATE_STRING.substring(6,8);
					
					var CONSENT_DATE_STRING = '' + CONSENT_END_DATE.v;
					var ConsentEndYear = CONSENT_DATE_STRING.substring(0,4);
					var ConsentEndMonth = CONSENT_DATE_STRING.substring(4,6);
					var ConsentEndDay = CONSENT_DATE_STRING.substring(6,8);
					
					//Definition of the string for the case creation
					var CaseCreationString = '<![CDATA[$XMLCoding<BizAgiWSParam><domain>domain</domain><userName>WebService</userName> <Cases><Case><Process>UpdateClientsConsents</Process><Entities><M_UCP_UpdateConsentReq><dConsentDate>$ConsentDate</dConsentDate><dConsentEndDate>$EndConsentDate</dConsentEndDate><kmClient>$CLIENTSTRING</kmClient><kpIncomingChannel><sCode>C_DB</sCode></kpIncomingChannel><xConsentsModification><M_ConsentStatus><bNewValue>$CONSENT_VALUE</bNewValue><pConsentType><sCode>M</sCode></pConsentType></M_ConsentStatus></xConsentsModification></M_UCP_UpdateConsentReq></Entities></Case></Cases></BizAgiWSParam>]]>';
					CaseCreationString = CaseCreationString.replace('$ConsentDate', ConsentYear+'-'+ConsentMonth+'-'+ConsentDay+'T00:00:00.000');
					CaseCreationString = CaseCreationString.replace('$EndConsentDate', ConsentEndYear+'-'+ConsentEndMonth+'-'+ConsentEndDay+'T00:00:00.000');
					CaseCreationString = CaseCreationString.replace('$CONSENT_VALUE', '');
					CaseCreationString = CaseCreationString.replace('$CLIENTSTRING',ClientString);
					CaseCreationString = CaseCreationString.replace('$XMLCoding',XMLCoding);
					
					var args = {casesInfo: CaseCreationString};

					var properties = {
							'ClientArguments':args,
							'isCase': true,
							'RowNumber': InitRow,
							'retry': 0,
							'client': clientWE,
							'callback': callback,
							'LogsFiles' : LogsFiles
					}
					//Call the create cases web service with the previously generated string				
					OneSOAPCall(properties);
					
					//Increase the requested counters
					RequestedClients++;
					RequestedCases++;
					Log(LogsFiles, 'debug','Requested Clients: ' + RequestedClients + '\tRequested Cases: ' + RequestedCases );
					
				}
				
				InitRow++;
			}
			else
				empty = true;
		}
	});});
	Log(LogsFiles, 'debug','Ended cycle');
};

function OneSOAPCall(properties) {
	var ClientArguments = properties.ClientArguments;
	var isCase = properties.isCase;
	var client = properties.client;
	var LogsFiles = properties.LogsFiles;
	
	Log(LogsFiles, 'debug',(isCase?'Case':'Entity') + ' is going to be created');
	Log(LogsFiles, 'debug','Arguments generated: ' + JSON.stringify(ClientArguments));
	
	if(isCase)
	{
	//var client = clientWE;
		client.createCasesAsString(ClientArguments, function(err,result){GetClientResponse(LogsFiles, err, result,properties)});
	}
	else
	{
	//var client = clientEM;
		client.saveEntityAsString(ClientArguments, function(err,result){GetClientResponse(LogsFiles, err, result,properties)});
	}
};

function CleanText(TextVar)
{
	TextVar += "";
	TextVar = TextVar.replace(/&/g,'&amp;');
	TextVar = TextVar.replace(/</g,'&lt;');
	TextVar = TextVar.replace(/>/g,'&gt;');
	//console.log(TextVar);
	return TextVar;
}

function GetClientResponse(LogsFiles, err, result, properties)
{
	var ClientArguments = properties.ClientArguments;
	var isCase = properties.isCase;
	var RowNumber = properties.RowNumber;
	var retry = properties.retry;
	var client = properties.client;
	var ReturnVar = null;
	var callback = properties.callback;
	if(err || !result ){
		console.log(LogsFiles, 'Error');
		console.log(LogsFiles, 'Retries: ' + retry);
		AppError = true;
		if(retry < 3)
		{
			Log(LogsFiles, 'error','Error');
			Log(LogsFiles, 'error','Retries: ' + retry);
			//Log('error',err);
			properties.retry = retry+1;
			setTimeout(OneSOAPCall, 60000, LogsFiles, properties, callback);
		}
		else
		{
			console.log(err);
			NumberOfErrorClients ++;
			if(isCase)
				NumberOfErrorCases ++;
			ErrorJson.push(RowNumber);
		}
	}
	else
	{
		var regexMessage = /<ErrorMessage>(.*)<\/ErrorMessage>/g;
		var regexDatabase = /Rollback not Allowed - Connection not found/g;
		var regex = /<M_Clients>(.*)<\/M_Clients>/g;
		if(isCase)
			regex = /<processId>0<\/processId>/g;
		var StringArguments = JSON.stringify(result);
		var found = false;
		var ErrorMessage = regexMessage.exec(StringArguments);
		var Error = regex.exec(StringArguments);
		var ErrorDatabas = regexDatabase.exec(StringArguments);
		if ( ErrorMessage !== null || (!isCase && Error == null) || (isCase && Error !== null) ) 
		{
			if(ErrorDatabas !== null && retry < 3)
			{
				AppError = true;
				Log(LogsFiles, 'error','Retry: ' + retry + ' -> ' + JSON.stringify(ClientArguments));
				//Log('error',result);
				properties.retry = retry+1;
				setTimeout(OneSOAPCall, 60000, LogsFiles, properties, callback);
			}
			else
			{
				NumberOfErrorClients ++;
				if(isCase)
					NumberOfErrorCases ++;
				ErrorJson.push(RowNumber);
				Log(LogsFiles, 'error',JSON.stringify(ClientArguments));
				Log(LogsFiles, 'error',result);
			}
		}
		else
		{				
			Log(LogsFiles, 'debug',result);
			Log(LogsFiles, 'debug',JSON.stringify(ClientArguments));
			
			//Increase the counters for clients generated
			NumberOfClients ++;		
			if(isCase)
				NumberOfCases ++;	
			Log(LogsFiles, 'debug','Current Created Clients: ' + NumberOfClients);
			if(isCase)
				Log(LogsFiles, 'debug','Current Created Cases: ' + NumberOfCases);
			ReturnVar = result;
		}	
	}
	
	if(NumberOfClients + NumberOfErrorClients == NumberExpected)
	{
		callback(LogsFiles,ErrorJson);
	}
}

function ExecuteError(LogsFiles, Error)
{
	Log(LogsFiles, 'info','Created Clients: ' + NumberOfClients);
	Log(LogsFiles, 'info','Error Clients: ' + NumberOfErrorClients);
	Log(LogsFiles, 'info','Created Cases: ' + NumberOfCases);
	Log(LogsFiles, 'info',"Rows With Errors: " + JSON.stringify(ErrorJson));
	Log(LogsFiles, 'info',"**********************	 	END		****************");
}
function Log(LogsFiles, level,message)
{
	switch(level) {
    case 'debug':
        LogsFiles.logFDebug.debug(message);
        break;
    case 'info':
        LogsFiles.logFDebug.info(message);
        LogsFiles.logFInfo.info(message);
        console.log(message);
        break;
    case 'error':
        LogsFiles.logFDebug.error(message);
        LogsFiles.logFInfo.error(message);
        LogsFiles.logFError.error(message);
        console.log(message);
        break;
    default:
        LogsFiles.logFDebug.debug(message);
	}
}
