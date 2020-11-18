

#----------- SCRIPT SETUP

        #  - Database server connection
        $sql_server = "192.168.1.124"
        $sql_database = "AdventureWorks2012"
        $sql_timeout = 60


        #  - Security data
        # SSPI security. domain validation
        #$sql_security = "Integrated Security=SSPI"
        # user an pass security
        $sql_security = "User Id=query;Password=querypass"

        #  - Select command
        #sql command. select command, view, procedure
        $sql_command = "
                SELECT [ProductID]
                ,[Name]
                ,[ProductNumber]
                ,[MakeFlag]
                ,[FinishedGoodsFlag]
                ,[Color]
                ,[SafetyStockLevel]
                ,[ReorderPoint]
                ,[StandardCost]
                ,[ListPrice]  
                FROM [AdventureWorks2012].[Production].[Product]        
        " 
        #direct select or view
        #$sql_command = "EXEC procedurename" #procedure execution. must return a select.


        #  - Export file configuration

        # temporary path for export file
        $Temporary_base_path = "c:\temp\" #ended with \

        #final destination for exported file
        $ExportedFile_Destination_path = "c:\temp\csv\"  #ended with \

        #name of the exported file
        $ExportedFile_Name = "Filename" #without extension

        #add timestamp to the filename 
        $AddTimesamp = 1   #[0|1]
    
        #debug mode on, gives you information about paths and the full coneccitions string.
        $DebugMode = 0 #[0|1]



#----------- END SCRIPT SETUP


#----------- PROCESS
try
{

    #create the connection string
    #connection string reference https://www.connectionstrings.com/    
    $connectionstringTemplate = 'Data Source={0};Initial Catalog={1};{2}'
    $connectionstring = [string]::Format($connectionstringTemplate, $sql_server, $sql_database, $sql_security)

    If ($DebugMode -eq 1) { Write-Host $connectionString }

    $now = Get-Date
    
    #set the final name of exported file
    $Exported_File_FinalName = $ExportedFile_Name
    If ($AddTimesamp -eq 1) 
    { 
        $timestamp = "{0:yyyyMMdd}_{0:HHmmss}" -f $now; 
        $Exported_File_FinalName = [string]::Format("{0}_{1}",$ExportedFile_Name,$timestamp)
    }
    $Exported_File_Final = [string]::Format("{0}{1}.csv",$ExportedFile_Destination_path,$Exported_File_FinalName) 
       
    If ($DebugMode -eq 1) { Write-Host $Exported_File_Final }
    
    #connect and exec command

    #create connection object
    $sql_connection = New-Object System.Data.SqlClient.SqlConnection
    $sql_connection.ConnectionString = $connectionstring  

    #command
    $command = New-Object System.Data.SqlClient.SqlCommand
    $command.CommandText = $sql_command
    $command.Connection = $sql_connection
    $command.CommandTimeout = $sql_timeout

    #adapter
    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SqlAdapter.SelectCommand = $command
    $DataSet = New-Object System.Data.DataSet

    #open - execute - close
    $sql_connection.Open()
    $SqlAdapter.Fill($DataSet)
    $sql_connection.Close()
    
    
    #export to csv
    #set temporary file
    $temporaryexportedfilepath = [string]::Format("{0}Temp_{1}.csv",$Temporary_base_path,$Exported_File_FinalName)
    If ($DebugMode -eq 1) { Write-Host $temporaryexportedfilepath }

    #export 2 csv. use ; for delimiter UTF 8 Encoding
    ($DataSet.Tables[0] | ConvertTo-Csv -Delimiter ";" -NoTypeInformation) -replace "`"", "" | Out-File  -encoding UTF8 -Force $temporaryexportedfilepath

    #move to end destination
    Move-Item -Path $temporaryexportedfilepath -Destination $Exported_File_Final 
    
}
catch [System.IO.IOException] 
{
    Write-Host "IO error"
    Write-Host $_
}
catch [System.Data.Common.DbException] 
{
    Write-Host "Database error"
    Write-Host $_
}
catch 
{
    Write-Host "An error occurred:"
    Write-Host $_
    Write-Host $_.ScriptStackTrace
}
finally 
{

}