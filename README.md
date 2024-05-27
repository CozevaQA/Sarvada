Chart List Export

1 . A folder will be created in downloads containing all the files 
	2. The files will be copied  to C://ChartListExport//customer_id wise 
	3. Reports will be generated in directory provided by user  

Steps : 

1. Create a report folder ChartList Reports
2. Run the script 
3. In the UI , it will start by applying date as date_from = 5 days from current date 
date_to = current date
4.. If the number of entries is less , it will increase the gap to 10 days 
if the number of entries is more than required , it will decrease the gap i.e. date_from will become 2 days from current date and date_to will remain the same
5. Reports will be downloaded a new timestamp folder in Downloads
6. It will be copied to ChartList Exports , based on customer id 
7. The paths will be passed to another function that will generate the report

