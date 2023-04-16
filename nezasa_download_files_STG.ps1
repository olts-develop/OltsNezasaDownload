# pwsh.exe -File nezasa_download_files_STG.ps1 -AGENCY xxxxxxx
# # Script requires Powershell 7, xmlstartlet and jsontoxml-cli-windows-amd64.exe to be in the PATH
# choco install -y powershell-core xmlstarlet
# # Run once to create secure credential file
# Get-Credential | EXPORT-CLIXML "SecureCredentialsSTG.xml"
# Get-Credential | EXPORT-CLIXML "SecureCredentialsPROD.xml"

param (
	$FILTERAFTER = "2022-01-01",
	$NEZASAHOST = "api.stg.nezasa.com",
	$NEZASAENV = "STG",
	$AGENCY = "",
	$PAGESIZE = 50,
	$BOOKINGSVERSIONNR = "v1.13",
	$PLANNERVERSIONNR	= "v1",
	$BOOKINGSTATES = "BookingCompleted,CancellationCompleted,BookingChangeCompleted"
)

# $fetchAll = $true
$fetchAll = $false
$FlagError = $false
$NewMaxDateTimeString = ""
$CurrentMaxDateTimeString = ""
$MaxDateTimeFileName = "MaxDateTime.txt" # The file name is set below to include STG or PROD

# For testing purposes. Comment out in production.
# if (Test-Path "MaxDateTimeProd.txt") {
#   Remove-Item "MaxDateTimeProd.txt"
#   Write-Host "Remove MaxDateTimeProd.txt"
# }

# $SecPass = ConvertTo-SecureString $PASS -AsPlainText -Force
# $Cred = New-Object System.Management.Automation.PSCredential ($USER, $SecPass)

# -----------------------------------------------------------------------
# ------------------------------ FUNCTIONS ------------------------------

Function Log {
    param(
        [Parameter(Mandatory=$true)][String]$msg
    )
    
	Write-Host $msg
    Add-Content $LogFile $msg
}

function NezasaGetBookings {
	
	param (
        $NezasaGetBookingsPageNumber
    )
	
	Log "-"
	Log "NezasaGetBookings"
	Write-Output "NezasaGetBookings"
	Log "  NezasaGetBookingsPageNumber = $NezasaGetBookingsPageNumber"
	
	$bookingsBody = @{
		'filter[modified-after]' = $FILTERAFTER
 		'filter[booking-states]' = $BOOKINGSTATES
		'page[size]' = $PAGESIZE
		'page[number]' = $NezasaGetBookingsPageNumber
		'agency' = $AGENCY 
	}

	$bookingsParams = @{
		Method = "Get"
		Uri = "https://$NEZASAHOST/bookings/$BOOKINGSVERSIONNR"
		Authentication = "Basic"
		Credential = $Cred
		Body = $bookingsBody
	}
	
	Log "URI:  $($bookingsParams.Uri)"

	$bookingsObj = New-Object -TypeName psobject
	
	try {
		$bookingsObj = Invoke-RestMethod @bookingsParams
		Log "Success"
	} catch {
		Log "Error"
		
		$FlagError = $true
		$ErrorObj = New-Object -TypeName psobject
		$ErrorObj | Add-Member -MemberType NoteProperty -Name Error -Value true
		$ErrorObj | Add-Member -MemberType NoteProperty -Name ErrorObject -Value "NezasaGetBookings"
		$ErrorObj | Add-Member -MemberType NoteProperty -Name Params -Value $($bookingsParams  | ConvertTo-JSON -depth 11)
		$ErrorObj | Add-Member -MemberType NoteProperty -Name StatusCode -Value $_.Exception.Response.StatusCode
		$ErrorObj | Add-Member -MemberType NoteProperty -Name IsSuccessStatusCode -Value $_.Exception.Response.IsSuccessStatusCode
		$ErrorObj | Add-Member -MemberType NoteProperty -Name Message -Value $_.Exception.Message
		
		Log "NezasaGetBookingObjJson: $($ErrorObj  | ConvertTo-JSON -depth 11)"
	}
	
	Return $bookingsObj
}

function NezasaGetNrPages {	
	$NezasaGetNrPagesObj = NezasaGetBookings -NezasaGetBookingsPageNumber 1
	Return $NezasaGetNrPagesObj.meta.pages	
}

function NezasaGetBooking {
	
	param (
        [string]$NezasaGetBookingId
    )
	
	# $NezasaGetBookingId = "test$NezasaGetBookingId"
	Log "-"
	Log "NezasaGetBooking"
	Log "  NezasaGetBookingId = $NezasaGetBookingId"
	
	$bookingParams = @{
		Method = "Get"
		Uri = "https://$NEZASAHOST/bookings/$BOOKINGSVERSIONNR/$NezasaGetBookingId"
		Authentication = "Basic"
		Credential = $Cred
	}
	
	Log "URI:  $($bookingParams.Uri)"
	
	$bookingObj = New-Object -TypeName psobject
	
	try {
		$bookingObj = Invoke-RestMethod -cred $Cred @bookingParams
		Log "Success"
		
		$created = $bookingObj.data.attributes.created
		$modified = $bookingObj.data.attributes.modified
		$maxDate = if ($created -gt $modified) {$created} else {$modified}
		
		$createdString = $created.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")		
		$modifiedString = $modified.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")		
		# $maxDateString = $maxDate.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
		$maxDateString = $maxDate.ToString("yyyy-MM-ddTHH:mm")
		
		Log "created:  $createdString"
		Log "modified: $modifiedString"
		Log "max:      $maxDateString"		
		
		if ( $maxDateString ) {			
			if ( $maxDateString -gt $NewMaxDateTimeString ) {
				Set-Variable -Name NewMaxDateTimeString -Value $maxDateString -Scope Script
				Log "new max:  $NewMaxDateTimeString"
			}
		}

		# $bookingObj  | ConvertTo-JSON -depth 11 | Out-File -FilePath .\booking-$IdFromArray.json
		$filePath = "$FileDirectory\booking_${IdFromArray}.xml"
		Log "  filePath = $filePath"
		$bookingObj  | ConvertTo-JSON -depth 11 | jsontoxml-cli-windows-amd64.exe | xml fo | Out-File -FilePath $filePath
		
	} catch {
		Log "Error"
		
		$FlagError = $true
		$ErrorObj = New-Object -TypeName psobject
		$ErrorObj | Add-Member -MemberType NoteProperty -Name Error -Value true
		$ErrorObj | Add-Member -MemberType NoteProperty -Name ErrorObject -Value "NezasaGetBooking"
		$ErrorObj | Add-Member -MemberType NoteProperty -Name NezasaGetBookingId -Value $NezasaGetBookingId
		$ErrorObj | Add-Member -MemberType NoteProperty -Name Params -Value $($bookingParams  | ConvertTo-JSON -depth 11)
		$ErrorObj | Add-Member -MemberType NoteProperty -Name StatusCode -Value $_.Exception.Response.StatusCode
		$ErrorObj | Add-Member -MemberType NoteProperty -Name IsSuccessStatusCode -Value $_.Exception.Response.IsSuccessStatusCode
		$ErrorObj | Add-Member -MemberType NoteProperty -Name Message -Value $_.Exception.Message
		
		Log "NezasaGetBookingObjJson: $($ErrorObj  | ConvertTo-JSON -depth 11)"
		Log "Exception: $($_.Exception  | ConvertTo-JSON -depth 11)"
	}
}

function NezasaGetPlannerItinerary {
	
	param (
        [string]$ItineraryId,
		[string]$ItineraryType,
		[string]$Suffix
    )
	
	Log "-"
	Log "NezasaGetPlannerItinerary"
	Log "  ItineraryId = $ItineraryId"
	Log "  ItineraryType = $ItineraryType"
	
	$plannerParams = @{
		Method = "Get"
		Uri = "https://$NEZASAHOST/planner/$PLANNERVERSIONNR/itineraries/$ItineraryId$Suffix"
		Authentication = "Basic"
		Credential = $Cred
	}
	
	Log "URI:  $($plannerParams.Uri)"
	
	$plannerObj = New-Object -TypeName psobject

	try {
		$plannerObj = Invoke-RestMethod -cred $Cred @plannerParams
		Log "Success"
		
		# $plannerObj | ConvertTo-JSON -depth 20 | Out-File -FilePath .\itinerary-$IdFromArray.json
		$filePath = "$FileDirectory\${ItineraryType}_${IdFromArray}.xml"
		Log "  filePath = $filePath"
		$plannerObj | ConvertTo-JSON -depth 20 | jsontoxml-cli-windows-amd64.exe | xml fo | Out-File -FilePath $filePath
	} catch {
		Log "Error"
		
		$ErrorObj = New-Object -TypeName psobject
		$ErrorObj | Add-Member -MemberType NoteProperty -Name Error -Value true
		$ErrorObj | Add-Member -MemberType NoteProperty -Name ErrorObject -Value "NezasaGetPlannerItinerary"
		$ErrorObj | Add-Member -MemberType NoteProperty -Name ItineraryId -Value $ItineraryId
		$ErrorObj | Add-Member -MemberType NoteProperty -Name Params -Value $($plannerParams  | ConvertTo-JSON -depth 11)
		$ErrorObj | Add-Member -MemberType NoteProperty -Name StatusCode -Value $_.Exception.Response.StatusCode
		$ErrorObj | Add-Member -MemberType NoteProperty -Name IsSuccessStatusCode -Value $_.Exception.Response.IsSuccessStatusCode
		$ErrorObj | Add-Member -MemberType NoteProperty -Name Message -Value $_.Exception.Message
		
		Log "NezasaGetPlannerItinerary: $($ErrorObj  | ConvertTo-JSON -depth 11)"
	}
}

# ------------------------------ FUNCTIONS ------------------------------
# -----------------------------------------------------------------------

Write-Output "NEZASAENV = $($NEZASAENV.ToLower())"
$envString = switch ($NEZASAENV.ToLower()) {
	"s" {"stg"}
	"st" {"stg"}
	"stg" {"stg"}
	"stage" {"stg"}
	"staging" {"stg"}
	"p" {"prod"}
	"pr" {"prod"}
	"prd" {"prod"}
	"prod" {"prod"}
	"production" {"prod"}
	default {"error"}
}
Write-Output "envString = $envString"

if ( $envString -eq "error" ) {
	Write-Output "Start: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff")"
	Write-Output "Error: invalid environment value $($NEZASAENV.ToUpper())"
	Write-Output "Valid values are STG and PROD"
	Exit 1
}

# Need a directory to store the log files
$LogPath = "$(Get-Location)\$envString\log"
If( !(test-path -PathType container $LogPath) )
{
      New-Item -ItemType Directory -Path $LogPath
}

# Need a directory to store the downloaded JSON and converted XML files
$FileDirectory = "$(Get-Location)\$envString\files"
If( !(test-path -PathType container $FileDirectory) )
{
      New-Item -ItemType Directory -Path $FileDirectory
}

# Remove all log files older than one day	
Get-ChildItem –Path $LogPath -Recurse | Where-Object {($_.LastWriteTime -lt (Get-Date).AddDays(-10))} | Remove-Item

# Create the new log file with date and time in the file name
$LogFile = "$LogPath\log_$(Get-Date -Format "yyyy-MM-dd_HH_mm_ss").txt"
if ( Test-Path $LogFile ) {
	Write-Output "Delete log.txt file"
	Remove-Item $LogFile
}
# -----------------------------------------------------------------------
# Start the workflow

Log "------------------------------------------------------------------------------"
Log "Start: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff")"
Log "-"
Log "LogFile: $LogFile"
Log "-"
Log "Parameters"
Log "FILTERAFTER:       $FILTERAFTER"
Log "NEZASAHOST:        $NEZASAHOST"
Log "AGENCY:            $AGENCY"
Log "PAGESIZE:          $PAGESIZE"
Log "BOOKINGSVERSIONNR: $BOOKINGSVERSIONNR"
Log "PLANNERVERSIONNR:  $PLANNERVERSIONNR"
Log "NEZASAENV:         $($NEZASAENV.ToLower())"
Log "envString:         $envString"

$Cred = IMPORT-CLIXML "SecureCredentials$($NEZASAENV.ToUpper()).xml" # The file name is set to include STG or PROD

# Get the last value for the MaxDateTime String
$TextInfo = (Get-Culture).TextInfo
$MaxDateTimeFileName = "MaxDateTime$($TextInfo.ToTitleCase($envString.ToLower())).txt"
Log "MaxDateTimeFileName: $MaxDateTimeFileName"
Log "-"
if ( Test-Path -Path $MaxDateTimeFileName -PathType Leaf ) {
	$CurrentMaxDateTimeString = Get-Content -Path $MaxDateTimeFileName -TotalCount 1
	Log "CurrentMaxDateTimeString from file: ´$CurrentMaxDateTimeString´"
	if ( !($fetchAll) ) {
		if ( $CurrentMaxDateTimeString ) {
			# $FILTERAFTER = $CurrentMaxDateTimeString.Substring(0,20)
			$FILTERAFTER = $CurrentMaxDateTimeString
			Log "Change FILTERAFTER: ´$CurrentMaxDateTimeString´"
		}
	}
}

$NrPages = NezasaGetNrPages

Log "-"
Log "NrPages of bookings: $NrPages"

if ($NrPages -gt 0) {

	$IdArrayList = New-Object System.Collections.ArrayList($null)

	If ($NrPages -gt 0) {
		Log "-"
		Log "Loop through Pages: $NrPages"
		For($Page = 1 ; $Page -le $NrPages ; $Page++) {
			Log "-"
			Log "NezasaGetBookings Page: $Page"
			
			$GetBookingsObj = NezasaGetBookings -NezasaGetBookingsPageNumber $Page
			
			Log "Found items: $($GetBookingsOBJ.data.id.Count)"
			
			ForEach ($IdFromArray in $GetBookingsObj.data.id) {
					[void]($IdArrayList.Add($IdFromArray))
			}
		}	
	}

	If ( $IdArrayList.Count -gt 0 ) {
		Log "-"
		Log "Loop through IdArrayList: $($IdArrayList.Count)"
		ForEach ($IdFromArray in $IdArrayList) {
			Log "----------"
			Log "Get booking and planner data for ID: $IdFromArray"
			
			NezasaGetBooking -NezasaGetBookingId $IdFromArray
			# Log "NewMaxDateTimeString (after function call):      $NewMaxDateTimeString"
			NezasaGetPlannerItinerary -ItineraryId $IdFromArray -ItineraryType "itinerary" -Suffix ""
			
			# -Suffix "/accommodations?include=accommodation.amenities,accommodation.tags,accommodation.contactDetails,accommodation.location"
			NezasaGetPlannerItinerary -ItineraryId $IdFromArray -ItineraryType "accommodations" -Suffix "/accommodations?include=accommodation.contactDetails"
			
			# -Suffix "/activities?include=accommodation.amenities,accommodation.tags,accommodation.contactDetails,accommodation.location,activity.segments,activity.amenities,activity.serviceCategories,activity.tourAttributes,activity.infoSections,activity.tags,activity.pictures,activity.paxDetails,transfer.amenities,transfer.serviceCategory,transfer.tags,transfer.type,transfer.private,transfer.shortDesc,rentalCar.acriss,rentalCar.pictures,rentalCar.tags,itinerary.paxDetails,itinerary.allocatedPax,itinerary.customerContact,itinerary.agency,itinerary.cancellation,itinerary.tourAttributes,itinerary.infoSections,module.tourAttributes,module.infoSections"
			NezasaGetPlannerItinerary -ItineraryId $IdFromArray -ItineraryType "activities" -Suffix "/activities?include=activity.segments,activity.amenities,activity.serviceCategories,activity.tourAttributes,activity.infoSections,activity.tags,activity.paxDetails,transfer.amenities,transfer.serviceCategory,transfer.tags,transfer.type,transfer.private,transfer.shortDesc,itinerary.paxDetails,itinerary.allocatedPax,itinerary.customerContact,itinerary.agency,itinerary.cancellation,itinerary.tourAttributes,itinerary.infoSections,module.tourAttributes,module.infoSections"
			
			# -Suffix "/rental-cars?include=accommodation.amenities,accommodation.tags,accommodation.contactDetails,accommodation.location,activity.segments,activity.amenities,activity.serviceCategories,activity.tourAttributes,activity.infoSections,activity.tags,activity.pictures,activity.paxDetails,transfer.amenities,transfer.serviceCategory,transfer.tags,transfer.type,transfer.private,transfer.shortDesc,rentalCar.acriss,rentalCar.pictures,rentalCar.tags,itinerary.paxDetails,itinerary.allocatedPax,itinerary.customerContact,itinerary.agency,itinerary.cancellation,itinerary.tourAttributes,itinerary.infoSections,module.tourAttributes,module.infoSections"
			NezasaGetPlannerItinerary -ItineraryId $IdFromArray -ItineraryType "rental-cars" -Suffix "/rental-cars?include=itinerary.customerContact,itinerary.agency,itinerary.tourAttributes,itinerary.infoSections,module.tourAttributes"
			
			# -Suffix "/transfers?include=accommodation.amenities,accommodation.tags,accommodation.contactDetails,accommodation.location,activity.segments,activity.amenities,activity.serviceCategories,activity.tourAttributes,activity.infoSections,activity.tags,activity.pictures,activity.paxDetails,transfer.amenities,transfer.serviceCategory,transfer.tags,transfer.type,transfer.private,transfer.shortDesc,rentalCar.acriss,rentalCar.pictures,rentalCar.tags,itinerary.paxDetails,itinerary.allocatedPax,itinerary.customerContact,itinerary.agency,itinerary.cancellation,itinerary.tourAttributes,itinerary.infoSections,module.tourAttributes,module.infoSections"
			NezasaGetPlannerItinerary -ItineraryId $IdFromArray -ItineraryType "transfers" -Suffix "/transfers?include=transfer.amenities,transfer.serviceCategory,transfer.tags,transfer.type,transfer.private,transfer.shortDesc"

		
			
		}	
	}

	Log "----------"
	Log "NewMaxDateTimeString:     $NewMaxDateTimeString"
	Log "CurrentMaxDateTimeString: $CurrentMaxDateTimeString"

	if ( !($FlagError) ) {
		Log "-"
		Log "^^^^^^^^^^ SUCCESS ^^^^^^^^^^"
		if ( $NewMaxDateTimeString -gt $CurrentMaxDateTimeString ) {
			Log "CurrentMaxDateTimeString smaller than NewMaxDateTimeString"
			Log "NewMaxDateTimeString saved to file $MaxDateTimeFileName = $NewMaxDateTimeString"
			$NewMaxDateTimeString | Out-File -FilePath $MaxDateTimeFileName			
		} else {
			Log "CurrentMaxDateTimeString == NewMaxDateTimeString = $NewMaxDateTimeString"
		}
		Log "^^^^^^^^^^ SUCCESS ^^^^^^^^^^"
	} else {
		Log "-"
		Log "^^^^^^^^^^ ERROR ^^^^^^^^^^"
		Log "^^^^^^^^^^ ERROR ^^^^^^^^^^"
		Log "There was an ERROR in the processing of the script"
		Log "Please analyse the log file:"
			
		$fileNameFromLogFile = Split-Path $LogFile -Leaf
		Log "fileNameFromLogFile = $fileNameFromLogFile"
		$directoryFromLogFile = Split-Path -Path $LogFile
		Log "directoryFromLogFile = $directoryFromLogFile"
		$newFileNameFromLogFile = "$fileNameFromLogFile" + ".error"
		Rename-Item $LogFile $newFileNameFromLogFile
		$LogFile = "$directoryFromLogFile\$newFileNameFromLogFile"
		
		Log "Renamed log file to *.error"
		Log "New log file = $LogFile"
		Log "-"
		Log "The MaxDateTime is NOT saved when there is an error."
		Log "-"
		Log "^^^^^^^^^^ ERROR ^^^^^^^^^^"
		Log "^^^^^^^^^^ ERROR ^^^^^^^^^^"
	}

}

Log "-"
Log "End: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff")"
Log "------------------------------------------------------------------------------"
Log "-"
Log "-"
Log "-"

# https://virtuallysober.com/2018/01/04/introduction-to-powershell-rest-api-authentication/
# # Run once to create secure credential file
# Get-Credential –Credential (Get-Credential) | EXPORT-CLIXML "SecureCredentialsSTG.xml"
# # Run at the start of each script to import the credentials
# $Cred = IMPORT-CLIXML "SecureCredentialsSTG.xml"

# # Run once to create secure credential file
# Get-Credential –Credential (Get-Credential) | EXPORT-CLIXML "C:\SecureString\SecureCredentials.xml"
# # Run at the start of each script to import the credentials
# $Cred = IMPORT-CLIXML "C:\SecureString\SecureCredentials.xml"
# $RESTAPIUser = $Cred.UserName
# $RESTAPIPassword = $Cred.GetNetworkCredential().Password

# Read-Host -Prompt "Press any key to continue"
