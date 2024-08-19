# Retrieve AT data from postholer.com Appalachian Trail data

# Get the AT data and all the section links
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

return 99


# Find the table data for Way points, Shelters, Resupply, and Road Crossings
$tables = $WebResponse.ParsedHtml.body.getElementsByTagName('Table')
$Cells=$tables|?{$_.cells}
foreach ($table in $Cells) {
	$name = $table.innerText -split "`n"|select -first 1
	if ($name -match "Popular Resupply") { $resupply = $table.innerHTML -split "`n"}
	elseif ($name -match "Road Crossings") { $RoadCrossings = $table.innerHTML -split "`n"}
	elseif ($name -match "Shelters") { $Shelters = $table.innerHTML -split "`n" }
	elseif ($name -match "Waypoint Detail") { $WayPoints = $table.innerHTML -split "`n" }
}

# Parse the HTML tables into arrays of objects

# Way Point Headers
$headers = '"' + $(if ($WayPoints[4] -match ">(.*)</TD>") {$matches[1]} else {"Col1"})
$headers += "`",`"" + $(if ($WayPoints[5] -match ">(.*)</TD>") {$matches[1]} else {"Col1"})
$headers += "`",`"" + $(if ($WayPoints[6] -match ">(.*)</TD>") {$matches[1]} else {"Col1"})
$headers += "`",`"" + $(if ($WayPoints[7] -match ">(.*)</TD>") {$matches[1]} else {"Col1"})
$headers += "`",`"" + $(if ($WayPoints[8] -match ">(.*)</TD>") {$matches[1]} else {"Col1"})
$headers += "`",`"" + $(if ($WayPoints[9] -match ">(.*)</TD>") {$matches[1]} else {"Col1"})
$headers += "`",`"" + $(if ($WayPoints[10] -match ">(.*)</TD>") {$matches[1]} else {"Col1"})
$headers += "`",`"" + $(if ($WayPoints[11] -match ">(.*)</TD>") {$matches[1]} else {"Col1"})
$headers += "`",`"" + $(if ($WayPoints[12] -match ">(.*)</TD>") {$matches[1]} else {"Col1"})
$headers +=  "`",`"Link`"`",`"index`""
$arrWayPoints = @()
$arrWayPoints += $headers 

# Shelters Headers
$headers = '"' + $(if ($Shelters[4] -match ">(.*)</TD>") {$matches[1]} else {"Col1"})
$headers += "`",`"" + $(if ($Shelters[5] -match ">(.*)</TD>") {$matches[1]} else {"col2"})
$headers += "`",`"" + $(if ($Shelters[6] -match ">(.*)</TD>") {$matches[1]} else {"Col3"}) + '"'
$headers +=  "`",`"Link`"`",`"index`""
$arrShelters = @()
$arrShelters += $headers 

# Resupply Headers
$headers = '"' + $(if ($Resupply[4] -match ">(.*)</TD>") {$matches[1]} else {"Col1"})
$headers += "`",`"" + $(if ($Resupply[5] -match ">(.*)</TD>") {$matches[1]} else {"col2"})
$headers += "`",`"" + $(if ($Resupply[6] -match ">(.*)</TD>") {$matches[1]} else {"Col3"}) + '"'
$headers += "`",`"" + $(if ($Resupply[7] -match ">(.*)</TD>") {$matches[1]} else {"Col4"}) + '"'
$headers +=  "`",`"Link`"`",`"index`""
$arrResupply = @()
$arrResupply += $headers

# Road Crossings Headers
$headers = "`"Mile`",`"Elev`",`"Type`",`"Name`",`"Link`",`"index`""
$arrRoadCrossings = @()
$arrRoadCrossings += $headers

# Get data from each section.  This includes trail data and shelters
foreach ($rec in $SectionData) {
	$url = $rec.link
	$index = $rec.index
	$fullLink = $rec.FullLink
	$mile = $rec.mile
	$section = $rec.section
	$description = $rec.Decription
	"$index - $url - $fullLink"
	if (!($linkAT -eq $fullLink)) {
		$WebResponse = Invoke-WebRequest $fullLink

		# Find the table data for Way points, Shelters, Resupply, and Road Crossings
		$tables = $WebResponse.ParsedHtml.body.getElementsByTagName('Table')
		$Cells=$tables|?{$_.cells}
		foreach ($table in $Cells) {
			$name = $table.innerText -split "`n"|select -first 1
			if ($name -match "Popular Resupply") { $resupply = $table.innerHTML -split "`n"}
			elseif ($name -match "Road Crossings") { $RoadCrossings = $table.innerHTML -split "`n"}
			elseif ($name -match "Shelters") { $Shelters = $table.innerHTML -split "`n" }
			elseif ($name -match "Waypoint Detail") { $WayPoints = $table.innerHTML -split "`n" }
		}
		#continue
	}
	
	# Way Point details
	$ct = 0
	$hashLine = @{}
	foreach ($line in ($WayPoints[13..($WayPoints.length)])) {
		if ($line.trim() -match "<TD class=alignRight>(.*)</TD>") {
			$ct++
			$hashLine[$ct] = $matches[1]
		} elseif ($line.Trim() -match "<TD class=alignLeft .*><A href=`"(.*)`" target=_blank>(.*)</A></TD>") {
			$link = $Matches[1]
			$descr = $matches[2]
		} elseif ($line.Trim() -match "^<TD class=alignLeft style=`"WIDTH: 25%; PADDING-LEFT: 7px`">(.*)</TD></TR>") {
			$comment = $matches[1]
			$arrWayPoints += "`""+$descr+"`",`""+$hashLine.1+"`",`""+$hashLine.2+"`",`""+$hashLine.3+"`",`""+$hashLine.4+"`",`""+$hashLine.5+"`",`""+$hashLine.6+"`",`""+$hashLine.7+"`",`""+$comment+"`",`""+$link+"`",`""+$index+"`"" 
			#$hashLine = @{}
			$ct = 0
		}
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
			$arrShelters +=  "`"$descr`",`"$elevation`",`"$mile`",`"$link`",`"$index`""
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
			$arrResupply += "`""+$hashLine.1+"`",`""+$hashLine.2+"`",`""+$hashLine.3+"`",`""+$location+"`",`""+$link+"`",`""+$index+"`""
			$hashLine = @{}
			$ct = 0
		}
	}
	$arrResupply += '"DFT is straight line distance from trail departure point.",,,,,'
		
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
			$arrRoadCrossings += "`""+$hashLine.1+"`",`""+$hashLine.2+"`",`""+$hashLine.3+"`",`""+$location+"`",`""+$link+"`",`""+$index+"`""
			$ct = 0
		}
	}

}

$arrWayPoints | out-file .\WayPoints.csv -Encoding utf8
$arrShelters | out-file .\Shelters.csv -Encoding utf8
$arrResupply | out-file .\Resupply.csv -Encoding utf8
$arrRoadCrossings | out-file .\RoadCrossings.csv -Encoding utf8
