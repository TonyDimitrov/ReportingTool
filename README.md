			Export tool manual (Automation report sending)

The tool (application) is design to send report to a person/s over email or over ftp.
Currently the tool can generate CSV or EXCEL format file.

	How the tool works (basics in steps)
	
1. We need to generate a store procedure (report)
	
2. Tool is executing store procedure and from result (output) CSV or EXCEL file is generated .

3. Generated file is sent to client/s over email or ftp.


	How the tool works (technical details in steps)	
	
1. We need to generate a store procedure
	At SQL server 10.0.17.20 there are databases SendaboxUtility and SendaboxUtility_UAT we create a store procedure on this database depending we want to test or not.
	- Store procedure should return standart SELECT of columns which we want to include in report.
	
	After we create a store procedure we have to tell the tool to use this store procedure and also specify few more configurations like type of file
	we want to be generated, report recipients and so. 
	To do this we use AT SQL server 10.0.17.20, databases SendaboxUtility and SendaboxUtility_UAT table EXTRACTION_REPORT to set all configurations.
	EXTRACTION_REPORT columns:
		id_extraction_data,
		extraction_storedprocedure,
		extraction_email,
		email_subject,
		extraction_format,
		file_name,
		email_text,
		is_active,
		send_compressed,
		send_empty,
		send_on_business_days,
		send_on_week_days,
		send_on_month_beginning,
		send_on_month_end,
		ignore_from_date,
		ignore_to_date,
		send_by_protocol,
		ftp_host,
		ftp_port,
		ftp_username,
		ftp_password,
		ftp_remote_folder,
	
	Filing the columns with examples:
	
	extraction_storedprocedure => dbo.Export_Daily_Shipments,
	extraction_email => tony.dimitrov@supernova-factory.com,alda.maglia@italmondo.com (email recipients coma separated),
	email_subject => Estrazione spedizioni sendabox al {DATE},
	extraction_format => CSV or XLSX,
	file_name => Export_Daily_Shipments.csv or Export_Daily_Shipments.xlsx, (mind the file extention for CSV and XLSX)
	email_text => Si trasmette in allegato il report delle spedizioni Sendabox relative all'anno corrente,
	is_active => 1 or 0 (only if report is active value 1 is put it will be sent we do not have to delete reports we can set value to 0 and report will be ignored),
	send_compressed => 1 or 0 (if we want file to be send compressed as zip we put value 1 if not put value 0),
	send_empty => 1 or 0 (if store procedure do not return results but we still want to send empty report put value to 1),
	
		Next few columns are specifying when report want to be sent. Please read below (4 Reporting tool filter explanation)
		send_on_business_days,
		send_on_week_days,
		send_on_month_beginning,
		send_on_month_end,
		ignore_from_date,
		ignore_to_date,
	
    Othrer columns settings:

		send_by_protocol => 1 0r 2 (if we want to be sent by email we put 1 if we want by ftp we put 2)
		ftp_host => ftp://ftp.italmondo.com (if we want to send by ftp we need to provide ftp parameters)
		ftp_port => 21 (if we want to send by ftp we need to provide ftp parameters),
		ftp_username => DeliveryNow, (if we want to sent by ftp we need to provide ftp parameters)
		ftp_password => some-passssss (if we want to sent by ftp we need to provide ftp parameters),
		ftp_remote_folder => FUORI_ZONA/OUT (if we want to sent by ftp we need to provide ftp parameters),
	
2. From the store procedure result(output) CSV or EXCEL file is generated.
	Depending on the settings in table EXTRACTION_REPORT tool will execute store procedure and generate file (CSV or XLSX). The return SELECT columns names will be used as headers to file.
	
3. Generated file is sent to client/s over email or ftp.
	Depending on settings in table EXTRACTION_REPORT report will be sent via email or ftp, compressed(zip) or not.
	There also number of options when report need to be sent (read additional info Reporting tool filter explanation).	
	
	
Tool(project is at tfs: http://109.168.96.231:8080/tfs/sendaboxsql/SendAboxSQL/Team%20SendAboxSQL project ReportExtraction.
Tool is set as a job and is run daily. (But you can specify when report to be sent)	
	
	
	4 Reporting tool filter explanation
		
Filters setting could be found in: [dbo].[EXTRACTION_REPORT] table.

If all filters are null, report will be send always.

1. Filter: send_on_business_days 
- Report will be sent if the current day is from monday to friday (inclusive) no metter of the country calendar. As in some countries working days may vary.

2.Filter: send_on_week_days
- D1 - Monday, D2 - Tuesday, D3 - Wednesday, D4 - Thursday, D5 - Friday, D6 - Saturday, D7 - Sunday,
any value e. g. D1 or d1 (case insensitive) set report to be send every week on particular day. More than one value can be set with coma separator:
d1,d2,d7 => Monday, Tuesday and Sunday report will be sent.

 
3. Filter: send_on_month_beginning send report in begining of every month.

4. Filter: send_on_month_end send report in end of every month.

5. ignore_from_date AND ignore_to_date report is not send between specified dates


PS:

Filters from 1 to 4 add/extend date or days when report is sent can be consider as OR operator.
e.g. send_on_business_days = true and  send_on_week_days = 'd6,d7' => every day report will be sent.

Filter 5 has oposite action and will everride all above filters and reports between specified dates will not be sent.
 Could be used in summer when we know all people will take two weeks holiday in August for example. 