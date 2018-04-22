# BizagiInitLoad

This project exposes a general nodejs method to read from an excel file and generate either entities of cases from each row of the excel sheet. 

## Things that must be configured:

Set the name of the name of the server with the location of the web application and the name of the project

```
	var ServerName = '10.10.10.10';
	var ProjectName = 'PROJECT_NAME';

```
  
Create one attribute per each column from your excel sheet

```
			ATTRIBUTE_1 = worksheet['A' + InitRow];
      
```
      
If required validate that the column has value

```
			if(ATTRIBUTE_4 != undefined && ATTRIBUTE_4.v != null && ATTRIBUTE_4.v != '') {log.debug('ATTRIBUTE_4 '+ATTRIBUTE_4.v);foundOne = true;}
      
```
      
Then define an xml for the element you want to create, e.g., if you want to create a client with three attributes:

```
var ClientString = '<M_Clients><sATTRIBUTE_3>$ATTRIBUTE_3</ATTRIBUTE_3><kpTypeofClient><sCode>C</sCode></kpTypeofClient><sATTRIBUTE_2><scode>$ATTRIBUTE_2</scode></sATTRIBUTE_2><ATTRIBUTE_1><sCode>$ATTRIBUTE_1</sCode></ATTRIBUTE_1></M_Clients>';

```

Then, once created the client you can modify the code to call the webservice you want using either the WorkflowEngine client (*clientWE*) of el cliente Entity Manager (*clientEM*). In this sample code you will find that depending of the value of one of the attributes either a client would be created or a case.


