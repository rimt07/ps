#Enable script execution
#Set-ExecutionPolicy Unrestricted

$dbServer 	=	"DB_SERVER";
$dbUser 	=	"DB_USERNAME";
$dbPassword	=	"DB_USERPWD";
$dbName		= 	"DB_NAME";

# Load needed assemblies 
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SMO") | Out-Null; 
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SMOExtended")| Out-Null; 
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.ConnectionInfo")| Out-Null; 

# Return all user databases on a sql server 
function getDatabases 
{ 
	param ($sql_server); 
	$databases = $sql_server.Databases | Where-Object {$_.IsSystemObject -eq $false}; 
	return $databases; 
} 
 
# Get all schemata in a database 
function getDatabaseSchemata 
{ 
	param ($sql_server, $database); 
	$db_name = $database.Name; 
	$schemata = $sql_server.Databases[$db_name].Schemas; 
	return $schemata; 
} 
 
# Get all tables in a database 
function getDatabaseTables 
{ 
	param ($sql_server, $database); 
	$db_name = $database.Name; 
	$tables = $sql_server.Databases[$db_name].Tables | Where-Object {$_.IsSystemObject -eq $false}; 
	return $tables; 
} 
 
# Get all stored procedures in a database 
function getDatabaseStoredProcedures 
{ 
	param ($sql_server, $database); 
	$db_name = $database.Name; 
	$procs = $sql_server.Databases[$db_name].StoredProcedures | Where-Object {$_.IsSystemObject -eq $false}; 
	return $procs; 
} 
 
# Get all user defined functions in a database 
function getDatabaseFunctions 
{ 
	param ($sql_server, $database); 
	$db_name = $database.Name; 
	$functions = $sql_server.Databases[$db_name].UserDefinedFunctions | Where-Object {$_.IsSystemObject -eq $false}; 
	return $functions; 
} 
 
# Get all views in a database 
function getDatabaseViews 
{ 
	param ($sql_server, $database); 
	$db_name = $database.Name; 
	$views = $sql_server.Databases[$db_name].Views | Where-Object {$_.IsSystemObject -eq $false}; 
	return $views; 
} 
 
# Get all table triggers in a database 
function getDatabaseTriggers 
{ 
	param ($sql_server, $database); 
	$db_name = $database.Name; 
	$tables = $sql_server.Databases[$db_name].Tables | Where-Object {$_.IsSystemObject -eq $false}; 
	$triggers = $null; 
	foreach($table in $tables) 
	{ 
		$triggers += $table.Triggers; 
	} 
	return $triggers; 
} 
 

# This function buils a list of links for database object types 
function buildLinkList 
{ 
	param ($array, $path);
	
	#Write-Host $array
	 
	$output = "<ul>"; 
	
	$outputSchema += "<li><b>Schemas:</b><ul>";
	$outputTrigger += "<li><b>Triggers:</b><ul>";
	$outputTable += "<li><b>Tables:</b><ul>";
	$outputStored += "<li><b>Stored Procedures:</b><ul>";
	$outputObjects += "<li><b>Objects:</b><ul>";
	
	foreach($item in $array) 
	{ 
		if($item.IsSystemObject -eq $false) # Exclude system objects 
		{     
			if([string]$item.GetType() -eq "Microsoft.SqlServer.Management.Smo.Schema") 
			{ 
			   $outputSchema += "`n<li>$item</li>"; 
			} 
			elseif([string]$item.GetType() -eq "Microsoft.SqlServer.Management.Smo.Trigger") 
			{ 
				$outputTrigger += "`n<li>$item</li>"; 
			} 
			elseif([string]$item.GetType() -eq "Microsoft.SqlServer.Management.Smo.Table") 
			{ 
			   $outputTable += "`n<li>$item</li>"; 
			} 
			elseif([string]$item.GetType() -eq "Microsoft.SqlServer.Management.Smo.StoredProcedure") 
			{ 
			   $outputStored += "`n<li>$item</li>"; 
			} 
			else
			{ 
			  $outputObjects += "`n<li>$item</li>"; 
			} 
		} 
	} 
	
	$output  += $outputSchema +"</ul></li>" + $outputTable +"</ul></li>" + $outputStored +"</ul></li>" + $outputObjects +"</ul></li>";
	
	$output += "</ul>"; 
	return $output; 
} 
 
# Return the DDL for a given database object 
function getObjectDefinition 
{ 
	param ($item); 
	$definition = ""; 
	# Schemas don't like our scripting options 
	if([string]$item.GetType() -eq "Microsoft.SqlServer.Management.Smo.Schema") 
	{ 
		$definition = $item.Script(); 
	} 
	else 
	{ 
		$options = New-Object ('Microsoft.SqlServer.Management.Smo.ScriptingOptions'); 
		$options.DriAll = $true; 
		$options.Indexes = $true; 
		$definition = $item.Script($options); 
	} 
	return "$definition"; 
} 
 
# This function will get the comments on objects 
# MS calls these MS_Descriptionn when you add them through SSMS 
function getDescriptionExtendedProperty 
{ 
	param ($item); 
	$description = "Empty."; 
	foreach($property in $item.ExtendedProperties) 
	{ 
		if($property.Name -eq "MS_Description") 
		{ 
			$description = $property.Value; 
		} 
		
		
	} 
	return $description; 
} 
 
# Gets the parameters for a Stored Procedure 
function getProcParameterTable 
{ 
	param ($proc); 
	$proc_params = $proc.Parameters; 
	$prms = $proc_params | ConvertTo-Html -Fragment -Property Name, DataType, DefaultValue, IsOutputParameter; 
	return $prms; 
} 
 
# Returns a html table of column details for a db table 
function getTableColumnTable 
{ 
	param ($table); 
	$table_columns = $table.Columns; 
	$objs = @(); 
	foreach($column in $table_columns) 
	{ 
		$obj = New-Object -TypeName Object; 
		$description = getDescriptionExtendedProperty $column; 
		Add-Member -Name "Name" -MemberType NoteProperty -Value $column.Name -InputObject $obj; 
		Add-Member -Name "DataType" -MemberType NoteProperty -Value $column.DataType -InputObject $obj; 
		#Add-Member -Name "Default" -MemberType NoteProperty -Value $column.Default -InputObject $obj; 
		Add-Member -Name "Identity" -MemberType NoteProperty -Value $column.Identity -InputObject $obj; 
		Add-Member -Name "PK" -MemberType NoteProperty -Value $column.InPrimaryKey -InputObject $obj; 
		Add-Member -Name "FK" -MemberType NoteProperty -Value $column.IsForeignKey -InputObject $obj; 
		Add-Member -Name "Description" -MemberType NoteProperty -Value $description -InputObject $obj; 
		$objs = $objs + $obj; 
	} 
	$cols = $objs | ConvertTo-Html -Fragment -Property Name, DataType, Identity, PK, FK, Description; 
	return $cols; 
} 
 
# Returns a html table containing trigger details 
function getTriggerDetailsTable 
{ 
	param ($trigger); 
	$trigger_details = $trigger | ConvertTo-Html -Fragment -Property IsEnabled, CreateDate, DateLastModified, Delete, DeleteOrder, Insert, InsertOrder, Update, UpdateOrder; 
	return $trigger_details; 
} 
 
 
 # Simple to function to write html pages 
function writeHtmlPage 
{ 
	param ($title, $heading, $body, $filePath); 
	$html = "<html> 
			 <head> 
				 <title>$title</title> 

				 <style>
					pre { margin-left:50px;  max-width:800px; white-space: pre-wrap; white-space: -moz-pre-wrap; white-space: -o-pre-wrap; background: #D3D3D3; font-family:Tahoma;  font-size:8pt; border: 1px solid #bebab0;}
					code { width:750px; white-space: pre-wrap; white-space: -moz-pre-wrap; white-space: -o-pre-wrap; display: block; padding: 3em 0.5em 3em 0em; }
					
					body { font-family:Verdana;    font-size:10pt; }
					h1 { font-family:Arial;    font-size:12pt; }
					h2 { font-family:Arial;    font-size:18pt; }
					h3 { font-family:Arial;    font-size:14pt; }
					h4 { font-family:Arial;    font-size:14pt; }
					td, th { border:1px solid black; border-collapse:collapse; }
					th { color:white; background-color:black; }
					table, tr, td, th { padding: 2px; margin: 0px font-family:Verdana;  font-size:10pt;}
					table { margin-left:50px; /*background-color:#D3D3D3;*/ font-family:Verdana;  font-size:10pt;}
					
					blockquote { width:750px; white-space: pre-wrap; white-space: -moz-pre-wrap; white-space: -o-pre-wrap; background: #D3D3D3;  border-left: 10px solid black;  margin: 1.5em 10px 1.5em 50px;  padding: 0.5em 10px;  quotes: '\201C''\201D''\2018''\2019';}
					blockquote:before {  color: black;  content: open-quote;  font-size: 4em;  line-height: 0.1em;  margin-right: 0.25em;  vertical-align: -0.4em;}
					blockquote p {  display: inline;}
				</style>
			 </head>
			 
			 <body style='background-color:snow;font-family:Verdana;color:black;font-size:15px;'> 
				 <h1>$heading</h1> 
				$body 
			 </body> 
			 </html>";
	$html | Out-File -FilePath $filePath; 
} 


Function New-WordText {
	Param (
		[string]$Text,
		[int]$Size = 10,
		[string]$Style = 'Normal',
		[Microsoft.Office.Interop.Word.WdColor]$ForegroundColor = "wdColorAutomatic",
		[switch]$Bold,
		[switch]$Italic,
		[switch]$NoNewLine
	)  
	Try {
		$Selection.Style = $Style
	} Catch {
		Write-Warning "Style: `"$Style`" doesn't exist! Try another name."
		Break
	}
 
	If ($Style -notmatch 'Title|^Heading'){
		$Selection.Font.Size = $Size  
		If ($PSBoundParameters.ContainsKey('Bold')) {
			$Selection.Font.Bold = 1
		} Else {
			$Selection.Font.Bold = 0
		}
		If ($PSBoundParameters.ContainsKey('Italic')) {
			$Selection.Font.Italic = 1
		} Else {
			$Selection.Font.Italic = 0
		}          
		$Selection.Font.Color = $ForegroundColor
	}
 
	$Selection.TypeText($Text)
 
	If (-NOT $PSBoundParameters.ContainsKey('NoNewLine')) {
		$Selection.TypeParagraph()
	}
}

Function New-WordTable {
	[cmdletbinding(
		DefaultParameterSetName='Table'
	)]
	Param (
		[parameter()]
		[object]$WordObject,
		[parameter()]
		[object]$Object,
		[parameter()]
		[int]$Columns,
		[parameter()]
		[int]$Rows,
		[parameter(ParameterSetName='Table')]
		[switch]$AsTable,
		[parameter(ParameterSetName='List')]
		[switch]$AsList,
		[parameter()]
		[string]$TableStyle,
		[parameter()]
		[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]$TableBehavior = 'wdWord9TableBehavior',
		[parameter()]
		[Microsoft.Office.Interop.Word.WdAutoFitBehavior]$AutoFitBehavior = 'wdAutoFitContent'
	)
	#Specifying 0 index ensures we get accurate data from a single object
	$Properties = $Object[0].psobject.properties.name
	$Range = @($WordObject.Paragraphs)[-1].Range
	$Table = $WordObject.Tables.add(
	$WordObject.Range,$Rows,$Columns,$TableBehavior, $AutoFitBehavior)
 
	Switch ($PSCmdlet.ParameterSetName) {
		'Table' {
			If (-NOT $PSBoundParameters.ContainsKey('TableStyle')) {
				#$Table.Style = "Medium Shading 1 - Accent 1"
			}
			$c = 1
			$r = 1
			#Build header
			$Properties | ForEach {
				Write-Verbose "Adding $($_)"
				$Table.cell($r,$c).range.Bold=1
				$Table.cell($r,$c).range.text = $_
				$c++
			}  
			$c = 1    
			#Add Data
			For ($i=0; $i -lt (($Object | Measure-Object).Count); $i++) {
				$Properties | ForEach {
					$Table.cell(($i+2),$c).range.Bold=0
					$Table.cell(($i+2),$c).range.text = $Object[$i].$_
					$c++
				}
				$c = 1 
			}                 
		}
		'List' {
			If (-NOT $PSBoundParameters.ContainsKey('TableStyle')) {
				#$Table.Style = "Light Shading - Accent 1"
			}
			$c = 1
			$r = 1
			$Properties | ForEach {
				$Table.cell($r,$c).range.Bold=1
				$Table.cell($r,$c).range.text = $_
				$c++
				$Table.cell($r,$c).range.Bold=0
				$Table.cell($r,$c).range.text = $Object.$_
				$c--
				$r++
			}
		}
	}
}



# This function creates all the html pages for our database objects 
function createObjectTypePages 
{ 
	param ($objectName, $objectArray, $filePath, $db); 
	New-Item -Path $($filePath + $db.Name + "\$objectName") -ItemType directory -Force | Out-Null; 
	# Create index page for object type 
	$page = $filePath + $($db.Name) + "\index.html"; 
	$list = buildLinkList $objectArray ""; 
	  #writeHtmlPage $objectName $objectName $list $page; 
	$body= $list ;
	
	if($objectArray -eq $null) 
	{ 
		$list = "No $objectName in $db"; 
	} 
   
	# Individual object pages 
	if($objectArray.Count -gt 0) 
	{ 
		foreach ($item in $objectArray) 
		{ 
			if($item.IsSystemObject -eq $false) # Exclude system objects 
			{ 
			
				$title = "<h1>Database objects</h1>"; 
				$body += "<h2>$item</h2>";# + [string]$item.GetType(); 
				
				$description = getDescriptionExtendedProperty($item); 
			   
				$body += "<h4>Description</h4><blockquote><p>$description</p></blockquote>";
				#$body += "<h2>Description</h2>$description"; 
				$definition = getObjectDefinition $item; 
				
			   
				if([string]$item.GetType() -eq "Microsoft.SqlServer.Management.Smo.Schema") 
				{ 
					$body += ""; 
				} 
				else 
				{ 
					if(([string]$item.GetType() -eq "Microsoft.SqlServer.Management.Smo.StoredProcedure") -or ([string]$item.GetType() -eq "Microsoft.SqlServer.Management.Smo.UserDefinedFunction")) 
					{ 
						$proc_params = getProcParameterTable $item; 
						#$body += "<h3>Object Definition</h3><pre><code class='sql'>$definition</code></pre>"; 
						$body += "<h4>Parameters</h4>$proc_params"; 
					} 
					elseif([string]$item.GetType() -eq "Microsoft.SqlServer.Management.Smo.Table") 
					{ 
						$cols = getTableColumnTable $item; 
						#$body += "<h3>Object Definition</h3><pre><code class='sql'>$definition</code></pre>"; 
						$body += "<h4>Columns</h4>$cols"; 
					} 
					elseif([string]$item.GetType() -eq "Microsoft.SqlServer.Management.Smo.View") 
					{ 
						$cols = getTableColumnTable $item; 
						#$body += "<h3>Object Definition</h3><pre><code class='sql'>$definition</code></pre>"; 
						$body += "<h4>Columns</h4>$cols</pre>"; 
					} 
					elseif([string]$item.GetType() -eq "Microsoft.SqlServer.Management.Smo.Trigger") 
					{ 
						$trigger_details = getTriggerDetailsTable $item; 
						#$body += "<h3>Object Definition</h3><pre><code class='sql'>$definition</code></pre>"; 
						$body += "<h4>Details</h4>$trigger_details"; 
					}                     
				} 
			  
			} 
			
		}  
		 writeHtmlPage $title $title $body $page; 
	} 
} 
 
 
# Root directory where the html documentation will be generated 
$filePath = "c:\database_documentation\"; 
New-Item -Path $filePath -ItemType directory -Force | Out-Null; 

# sql server that hosts the databases we wish to document 
#$sql_server = New-Object Microsoft.SqlServer.Management.Smo.Server localhost\sqlexpress; 

$sql_server = new-object ('Microsoft.SqlServer.Management.Smo.Server') $dbServer  

$sql_server.ConnectionContext.LoginSecure = $false 
$sql_server.ConnectionContext.Login= $dbUser 
$sql_server.ConnectionContext.Password= $dbPassword 
$db = $sql_server.Databases.Item($dbName)

# IsSystemObject not returned by default so ask SMO for it 
$sql_server.SetDefaultInitFields([Microsoft.SqlServer.Management.SMO.Table], "IsSystemObject"); 
$sql_server.SetDefaultInitFields([Microsoft.SqlServer.Management.SMO.View], "IsSystemObject"); 
$sql_server.SetDefaultInitFields([Microsoft.SqlServer.Management.SMO.StoredProcedure], "IsSystemObject"); 
$sql_server.SetDefaultInitFields([Microsoft.SqlServer.Management.SMO.Trigger], "IsSystemObject"); 

 
# Get databases on our server 
#$databases = getDatabases $sql_server; 
 
#foreach ($db in $databases) 
##{ 
	Write-Host "Started documenting " $db.Name; 
	# Directory for each database to keep everything tidy 
	New-Item -Path $($filePath + $db.Name) -ItemType directory -Force | Out-Null; 

		 
	# Get schemata for the current database 
	$schemata = getDatabaseSchemata $sql_server $db; 
	#createObjectTypePages "Schemata" $schemata $filePath $db; 
	Write-Host "Documented schemata"; 
	
	# Get tables for the current database 
	$tables = getDatabaseTables $sql_server $db;
	#createObjectTypePages "Tables" $tables $filePath $db; 
	Write-Host "Documented tables"; 
	
	# Get views for the current database 
	$views = getDatabaseViews $sql_server $db; 
	#createObjectTypePages "Views" $views $filePath $db; 
	Write-Host "Documented views"; 
	
	# Get procs for the current database 
	$procs = getDatabaseStoredProcedures $sql_server $db; 
	#createObjectTypePages "Stored Procedures" $procs $filePath $db; 
	Write-Host "Documented stored procedures"; 
	
	# Get functions for the current database 
	$functions = getDatabaseFunctions $sql_server $db; 
	#createObjectTypePages "Functions" $functions $filePath $db; 
	Write-Host "Documented functions"; 
	
	# Get triggers for the current database 
	$triggers = getDatabaseTriggers $sql_server $db; 
	#createObjectTypePages "Triggers" $triggers $filePath $db; 
	Write-Host "Documented triggers"; 
	Write-Host "Finished documenting " $db.Name; 
	
	$all = $schemata + $tables + $views + $procs + $functions + $triggers  ;
	#Write-Host $all;
	createObjectTypePages "All" $all $filePath $db; 
	
 
	# make the powershell process switch the current directory.
	$oldwd = [Environment]::CurrentDirectory
	[Environment]::CurrentDirectory = $pwd
	 
	$html = $filePath + $db.Name + "\index.html";
	$docx = $filePath + $db.Name + "\"+ $db.Name + ".docx";
	[Environment]::CurrentDirectory = $oldwd
	 
	[ref]$SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type] 
	$word = New-Object -ComObject word.application 
	$word.visible = $false 
	 
	Write-Host  "Converting $html to $docx..." 
	$doc = $word.documents.open($html) 

	$Selection = $word.Selection

	###Create a table of revisions
	New-WordText -Text "Revisions" -Size 24 -Bold
	New-WordText -Text " " -Bold 

	#$obj = New-Object -TypeName Object; 
	#Add-Member -Name "Date" -MemberType NoteProperty -Value "value" -InputObject $obj; 
	#Add-Member -Name "Name" -MemberType NoteProperty -Value "value" -InputObject $obj;    
	#Add-Member -Name "Observations" -MemberType NoteProperty -Value "value" -InputObject $obj; 

	#New-WordTable -Object $obj -Columns 3 -Rows ($obj.Count+1) -AsTable -WordObject $Selection
	
	#$word.Selection.Start= $doc.Content.Start
	#$Selection = $word.Selection
	#$Selection.TypeParagraph()

	###Create a table of revisions
	New-WordText -Text "Index" -Size 24 -Bold
	New-WordText -Text " " -Bold 

	$range = $Selection.Range
	$toc = $doc.TablesOfContents.Add($range)
	$Selection.TypeParagraph()

	$doc.saveas([ref] $docx, [ref]$SaveFormat::wdFormatDocumentDefault) 

	##$word.Visible = $true
	$doc.close() 
	$word.Quit() 
	$word = $null 
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()

#}
