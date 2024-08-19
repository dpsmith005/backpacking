
<# Planner data
$URL = "https://www.postholer.com/planner/Appalachian-Trail/3"  	# Planner
$WebPlanner = Invoke-WebRequest $URL
$PlannerData = $WebPlanner.content -split "`n"
#id="planTable"
$tables = $WebPlanner.ParsedHtml.body.getElementsByTagName('Table')
$Cells=$tables|?{$_.cells}
$planTable = ($cells | ?{$_.id -eq "planTable"}).InnerHTML
#>
<#
 http://gis.postholer.com/services/reflect?service=WFS&typename=mileMarker&trail_id=3&types=1,8,,100,112,150&srs=EPSG:4269&bbox=-119.8,36.4,-118,38&outputformat=text/csvile
 http://gis.postholer.com/services/reflect?service=WFS&typename=wptByTrailType&trail_id=1&types=1,112&srs=EPSG:4269&bbox=-119.8,36.4,-118,38&outputformat=text/csv
 http://gis.postholer.com/services/reflect?
   service=WFS
   &typename=mileMarker
   &trail_id=3
   &types=1,8,,100,112,150
   &srs=EPSG:4269
   &bbox=-119.8,36.4,-118,38
   &outputformat=text/csvile ($ie.Busy -eq $true) {Start-Sleep -Seconds 3;}
&typename=wptByTrailType
#>

$x=Invoke-WebRequest 'https://www.postholer.com/planner/hikePlanner.php'
#$x=Invoke-WebRequest 'https://www.postholer.com/planner/Appalachian-Trail/3'
#$x.forms
<#
$lastTrailId = $x.ParsedHtml.getElementsByName("lastTrailId")
$($lastTrailId).value=3
$showAll=$x.ParsedHtml.getElementsByName("showAll")
$($showAll).checked=$true
$stMon = $x.ParsedHtml.getElementsByName("stMon")
$($stMon).value=3
$stDay = $x.ParsedHtml.getElementsByName("stDay")
$($stDay).value=20
$stYear = $x.ParsedHtml.getElementsByName("stYear")
$($stYear).value=2027
$params = @{showAll=$showAll;lastTrailId=$lastTrailId;stMon=$stMon;stDay=$stDay;stYear=$stYear}
$x=Invoke-WebRequest 'https://www.postholer.com/planner/hikePlanner.php' -Method POST -Body $Params
#$data = $x.content -split "`n"
#>

$data = get-content ".\ATPlanner.html"

$ct = 0
$header = ""
$flagHeader = $false
$flagDetail = $false
$lineDetail = ""
$arrDetail = @()
foreach ($line in $data) {
	$ct++
	
	# Find the header
	if ($line.Trim() -match "^<table id=`"planTable`" ") {
		$flagHeader = $true
		continue
	}
	if ($flagHeader -and $line.Trim() -match "^<td style=.*`">(.*)</td>$") {
		$header += $matches[1] + ","
		continue
	}
	if ($flagHeader -and $line -match "</tr>") {
		$flagHeader = $false
		$header = $header.Trim(",")
		$header = (($header.Replace("&nbsp","")).replace(";","")).replace("<br>","_")
		$arrDetail = $header + "`n"
		continue
	}
	
	# Finde the standard details
	if ($line.Trim() -match "^<td style=`"width: 25%; `"><strong>\d+ - (.*)</strong>.*</strong>") {
		$lineDetail = $matches[1]
		$flagDetail = $true
		continue
	}
	if ($line.Trim() -match "^<td style=`"width: 25%.*`">(.*)</td>$") {
		$lineDetail = $matches[1]
		$flagDetail = $true
		continue
	}
	if ($flagDetail -and $line.Trim() -match "^<td style=`"text-align:.*`">(.*)</td>$") {
		$value = $matches[1]
		if ($value -match "<br>") { 
			$value = "," + $value.split("\<")[0] 
		} 
		$lineDetail += "," + $value.Replace(",","") 
		continue
	}
	if ($flagDetail -and $line.Trim() -eq "</tr>") {
		$flagDetail = $false
		#$lineDetail = $lineDetail.Substring(0,($lineDetail.Length - 1))	#Drop last comma (,)
		$lineDetail = ($lineDetail.Replace("<strong>","")).Replace("</strong>","")	#Remove <strong> and </strong>
		$lineDetail = $lineDetail.Replace("</a>","")	#Remove </a>
		$arrDetail += $lineDetail + "`n"
		$lineDetail = ""
		continue
	}
	
	# Find the notes rows details
	#<tr id="notes_row_0" style="display: none; width: 100%;">
	#</tbody></table>
}

$arrDetail | out-file .\ATPlannerData.csv  -Encoding utf8