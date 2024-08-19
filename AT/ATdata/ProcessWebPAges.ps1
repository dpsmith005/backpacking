# Get the AT data and all the section links from postholer
$linkAT = "https://www.postholer.com/databook/Appalachian-Trail/3/0.0"
$WebResponse = Invoke-WebRequest $linkAT
$webData = $WebResponse.content -split "`n"
$flag = $false
$sections = @()
foreach ($line in $webData) {
	if ($line -match "metaTD") { $flag = $true }
	if ($flag) {
		$sections += $line
	}
	if ($line -match "</table>") { $flag = $false }
}
$SectionData = @()
$links = @()
foreach ($line in $sections) {
	#<option value="53.3"  >Mile 53.3-166.4, Section 2 Hiawassee GA</option>
	if ($line -match "\<option value=`"(.*)`"\s+\>(.*), (Section \d+) (.*)</option>" -or $line -match "\<option value=`"(.*)`"\s+selected\s+\>(.*), (Section \d+) (.*)</option>") {
		$link = $matches[1]
		$fullLink = "https://www.postholer.com/databook/Appalachian-Trail/3/" + $matches[1]
		$mile = $matches[2]
		$section = $matches[3]
		$index = ($section -split " ")[1]
		$desc = $matches[4]
		$SectionData += [pscustomobject]@{"section"=$section;"mile"=$mile;"Description"=$desc;"index"=$index;"link"=$link;"FullLink"=$fullLink}
		$links += "https://www.postholer.com/databook/Appalachian-Trail/3/$link"
	}
} 
$SectionData | Export-Csv -NoTypeInformation -Path .\SectionData.csv
$links | out-file .\links.txt

# Gathered AT data from Postholer.com as save to WebPages

# Read the AT data html files
$files = Get-ChildItem webpages\*.html

# Get section number for each file
$hashFiles = @{}
$files|%{if($_.name -match "Section (\d+) \w") {$hashFiles[[int]$matches[1]] = $_.Fullname} }

$hashShelters = @{}
$hashRoadCrossings = @{}
$hashResupply = @{}
$hashWayPoints = @{}

# loop through each trail section
foreach ($section in ($hashFiles.keys|sort)) {
	$file = $hashFiles.$section
	$file
	$index = $section
	$source=get-content $file -raw
	$HTML = New-Object -Com "HTMLFile"
	$HTML.IHTMLDocument2_write($Source)

	#$HTML.all.tags("code") | % InnerText

	# Find the table data for Way points, Shelters, Resupply, and Road Crossings
	$tables=$html.all.tags("table")
	$Cells=$tables|?{$_.cells}
	foreach ($table in $Cells) {
		$name = $table.innerText -split "`n"|select -first 1
		if ($name -match "Popular Resupply") { $resupply = $table.innerHTML -split "`n"}
		elseif ($name -match "Road Crossings") { $RoadCrossings = $table.innerHTML -split "`n"}
		elseif ($name -match "Shelters") { $Shelters = $table.innerHTML -split "`n" }
		elseif ($name -match "Waypoint Detail") { $WayPoints = $table.innerHTML -split "`n" }
	}
	
	if ($section -eq 1) {
		"Get header data"
		
		# Way Point Headers
		$headers = "`"Order"
		$headers += "`",`""+ $(if ($WayPoints[4] -match ">(.*)</TD>") {$matches[1]} else {"Col1"})
		$headers += "`",`"" + $(if ($WayPoints[5] -match ">(.*)</TD>") {$matches[1]} else {"Col2"})
		$headers += "`",`"" + $(if ($WayPoints[6] -match ">(.*)</TD>") {$matches[1]} else {"Col3"})
		$headers += "`",`"" + $(if ($WayPoints[7] -match ">(.*)</TD>") {$matches[1]} else {"Col4"})
		$headers += "`",`"" + $(if ($WayPoints[8] -match ">(.*)</TD>") {$matches[1]} else {"Col5"})
		$headers += "`",`"" + $(if ($WayPoints[9] -match ">(.*)</TD>") {$matches[1]} else {"Col6"})
		$headers += "`",`"" + $(if ($WayPoints[10] -match ">(.*)</TD>") {$matches[1]} else {"Col7"})
		$headers += "`",`"" + $(if ($WayPoints[11] -match ">(.*)</TD>") {$matches[1]} else {"Col8"})
		$headers += "`",`"" + $(if ($WayPoints[12] -match ">(.*)</TD>") {$matches[1]} else {"Col9"})
		$headers +=  "`",`"Link`",`"index`",`"Type`",`"Other`""
		$arrWayPoints = @()
		$arrWayPoints += $headers 

		# Shelters Headers
		$headers = '"' + $(if ($Shelters[4] -match ">(.*)</TD>") {$matches[1]} else {"Col1"})
		$headers += "`",`"" + $(if ($Shelters[5] -match ">(.*)</TD>") {$matches[1]} else {"col2"})
		$headers += "`",`"" + $(if ($Shelters[6] -match ">(.*)</TD>") {$matches[1]} else {"Col3"})  # + '"'
		$headers +=  "`",`"Link`",`"index`",`"Type`""
		$arrShelters = @()
		$arrShelters += $headers 

		# Resupply Headers
		$headers = '"' + $(if ($Resupply[4] -match ">(.*)</TD>") {$matches[1]} else {"Col1"})
		$headers += "`",`"" + $(if ($Resupply[5] -match ">(.*)</TD>") {$matches[1]} else {"col2"})
		$headers += "`",`"" + $(if ($Resupply[6] -match ">(.*)</TD>") {$matches[1]} else {"Col3"}) 
		$headers += "`",`"" + $(if ($Resupply[7] -match ">(.*)</TD>") {$matches[1]} else {"Col4"}) 
		$headers +=  "`",`"Link`",`"index`",`"Type`""
		$arrResupply = @()
		$arrResupply += $headers

		# Road Crossings Headers
		$headers = "`"Mile`",`"Elev`",`"Type`",`"Name`",`"Link`",`"index`",`"Type`""
		$arrRoadCrossings = @()
		$arrRoadCrossings += $headers
	}
	
	# Shelters
	foreach ($line in ($shelters[7..($Shelters.length)])) {
		if ($line.Trim() -match "^<TD class=alignLeft><A href=`"(.*)`" target=_blank>(.*)</A></TD>") {
			#$link = "https://www.postholer.com/" + $matches[1]
			$link = $matches[1]
			$descr = $matches[2]
			# 					<TD class=alignRight>3881</TD>
		} elseif ($line.Trim() -match "^<TD class=alignRight>(.*)</TD>$") {
			$elevation = $matches[1]
		} elseif ($line.Trim() -match "^<TD class=alignRight>(.*)</TD></TR>") {
			$mile = $matches[1]
			$arrShelters +=  "`"$descr`",`"$elevation`",`"$mile`",`"$link`",`"$index`",`"S`""
			$key = [double]$mile
			$hashShelters[$key] = "`"$descr`",`"$elevation`",`"$mile`",`"$link`",`"$index`",`"S`""
		}
	}
	
	# Resupply
	$ct = 0
	$hashLine = @{}
	foreach ($line in ($Resupply[8..($Resupply.length)])) {
		if ($line.trim() -match "<TD class=alignRight>(.*)</TD>") {
			$ct++
			$hashLine[$ct] = $matches[1]
		} elseif ($line.Trim() -match "<TD class=alignLeft><A href=`"(.*)`" target=_blank>(.*)</A></TD></TR>") {
			$link = $Matches[1]
			$location = $matches[2]
			$arrResupply += "`""+$hashLine.1+"`",`""+$hashLine.2+"`",`""+$hashLine.3+"`",`""+$location+"`",`""+$link+"`",`""+$index+"`",`"S`""
			$key = [double]$hashLine.1
			$hashResupply[$key] = "`""+$hashLine.1+"`",`""+$hashLine.2+"`",`""+$hashLine.3+"`",`""+$location+"`",`""+$link+"`",`""+$index+"`",`"S`""
			$hashLine = @{}
			$ct = 0
		}
	}
		
	# Road Crossings
	$ct = 0
	$hashLine = @{}
	foreach ($line in $RoadCrossings[8..($RoadCrossings.length)]) {
		if ($line.trim() -match "^<TD class=alignRight>(.*)</TD>$") {
			$ct++
			$hashLine[$ct] = $matches[1]
		} elseif ($line.trim() -match "^<TD class=alignLeft>(.*)</TD>$") {
			$ct++
			$hashLine[$ct] = $matches[1]
	##### START HERE - BELOWIS NOT WORKING #####
		} elseif ($line.Trim() -match "<td class=alignLeft><a href=`"(.*)`" (.*)>(.*)</A></TD></TR>") {
			$link = $Matches[1]
			$location = $matches[3]
			$arrRoadCrossings += "`""+$hashLine.1+"`",`""+$hashLine.2+"`",`""+$hashLine.3+"`",`""+$location+"`",`""+$link+"`",`""+$index+"`",`"S`""
			$key = [double]$hashLine.1
			$hashRoadCrossings[$key] =  "`""+$hashLine.1+"`",`""+$hashLine.2+"`",`""+$hashLine.3+"`",`""+$location+"`",`""+$link+"`",`""+$index+"`",`"S`""
			$ct = 0
		}
	}

	# Way Point details
	Remove-Variable mile
	$ct = 0
	$order = 0
	$hashLine = @{}
	foreach ($line in ($WayPoints[13..($WayPoints.length)])) {
		if ($line.trim() -match "<TD class=alignRight>(.*)</TD>") {
			$ct++
			$hashLine[$ct] = $matches[1]
		} elseif ($line.Trim() -match "<TD class=alignLeft .*><A href=`"(.*)`" target=_blank>(.*)</A></TD>") {
			$link = $Matches[1]
			$descr = $matches[2]
		} elseif ($line.Trim() -match "^<TD class=alignLeft style=`"WIDTH: 25%; PADDING-LEFT: 7px`">(.*)</TD></TR>") {
			$order++
			$comment = $matches[1]
			if (![string]::IsNullOrEmpty($hashline.1)) {$mile = $hashLine.1}
			$value = ""
			if ($hashShelters[$key]) {$value ="S "}
			if ($hashResupply[$key]) {$value += "R "}
			if ($hashRoadCrossings[$key]) {$value += "X"}
			$value = $value.Trim()
			$arrWayPoints += "`""+$order+"`",`""+$descr+"`",`""+$mile+"`",`""+$hashLine.2+"`",`""+$hashLine.3+"`",`""+$hashLine.4+"`",`""+$hashLine.5+"`",`""+$hashLine.6+"`",`""+$hashLine.7+"`",`""+$comment+"`",`""+$link+"`",`""+$index+"`",`"W`",`""+$value+"`"" 
			$key = [double]$mile
			$hashWayPoints[$key] = "`""+$order+"`",`""+$descr+"`",`""+$mile+"`",`""+$hashLine.2+"`",`""+$hashLine.3+"`",`""+$hashLine.4+"`",`""+$hashLine.5+"`",`""+$hashLine.6+"`",`""+$hashLine.7+"`",`""+$comment+"`",`""+$index+"`",`"W`",`""+$link+"`"" 
			#$hashLine = @{}
			$ct = 0
		}
	}
	
}

# Show Shelter, Resupply, RoadCrossing in WayPoints
foreach ($key in ($hashWayPoints.keys|sort)) {
	$value = ""
	if ($hashShelters[$key]) {$value ="S "}
	if ($hashResupply[$key]) {$value += "R "}
	if ($hashRoadCrossings[$key]) {$value += "X"}
	$value = $value.Trim()
	$hashWayPoints[$key] += ",`"$value`""   
}

# save as CSV files
$arrWayPoints | out-file .\WayPoints.csv -Encoding utf8
$arrShelters | out-file .\Shelters.csv -Encoding utf8
$arrResupply += '"DFT is straight line distance from trail departure point.",,,,,'
$arrResupply | out-file .\Resupply.csv -Encoding utf8
$arrRoadCrossings | out-file .\RoadCrossings.csv -Encoding utf8

# Combine CSV files into one XLXS file
$OutputFile = "C:\Users\dsmith14\OneDrive - WellSpan Health\Documents\GitHub\dps\at\ATdata\combined-data.xlsx"
Remove-Item $OutputFile -ea 0
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault
$excel            = New-Object -ComObject Excel.Application
#$excel.Visible    = $true

$workbook         = $excel.Workbooks.add()
$defaultsheet     = $workbook.Worksheets.Item(1)

foreach ($file in @('.\RoadCrossings.csv','.\Resupply.csv','.\Shelters.csv','.\SectionData.csv','.\WayPoints.csv')) {
	$csv = Get-ChildItem $file
	$csv
	$worksheet = $workbook.Worksheets.Add()
    $worksheet.Name = "$($csv.BaseName)"
    $TxtConnector = ("TEXT;" + $csv.fullname)
    $Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
    $query = $worksheet.QueryTables.item($Connector.name)
    $query.TextFileOtherDelimiter = ','
    $query.TextFileParseType  = 1
    $query.TextFileColumnDataTypes = ,1 * $worksheet.Cells.Columns.Count
    $query.AdjustColumnWidth = 1
    $query.Refresh() | Out-Null
    $query.Delete() | Out-Null
}

$defaultsheet.Delete() | Out-Null
$workbook.sheets.item("Sheet2").Delete()
$workbook.sheets.item("Sheet3").Delete()
$workbook.sheets.item("Sheet4").Delete()

#$workbook.SaveAs($OutputFile, $xlFixedFormat)
$excel.ActiveWorkbook.SaveAs("$OutputFile", $xlFixedFormat)
Start-Sleep -s 2
$workbook.Close()
$excel.Quit()
$excel = $null


<#
GIS data (Not working)
 http://gis.postholer.com/services/reflect?service=WFS&typename=wptByTrailType&trail_id=3&types=1,8&srs=EPSG:4269&bbox=-119.8,36.4,-118,38&outputformat=text/csv
        gis.postholer.com/services/reflect?service=WFS&typename=wptByTrailType&trail_id=3&types=1,8&srs=EPSG:4269&startmile=0&endmile=2100&outputformat=text/csv

#>
