(*
 * apertureExport
 * written by Jörn Stein <dev@justcoding.de>, 01/2017
 *
 * Usage: select all images that should be exported. Check the settings variables below, then run
 *
 * The script will export each master file into a subfolder YYYY/Projectname/CameraName (if the project
 * name begins with 4 digits) or "ProjectName/CameraName" (otherwise). Every edited version of the image
 * will be exported into "[YYYY/]ProjectName/_edited".
 * Optionally the rating, color rating, flag status and face names will be added as keywords.
 * Check the comments below for more information about what each setting does.
 *
 * During a run, some keywords may be added to the pictures in Aperture (see "addRatingsAndFacesAsKeywords"
 * below). Otherwise the Aperture library will be left untouched and the script can be run multiple times
 * if something does not go as expected.
 *
 * Note: sometimes Aperture cannot add IPTC data to the exported master file and will display a dialogue box
 * that the user has to confirm. It is possible that the script fails to export subsequent exports. However,
 * failed exports will be queued again for later processing, so no image should be missing at the end.
 *
 * Exporting 3000 photos at once is not a problem, maximum tried export size was 8000 pictures. YMMV.
 *)

-- Choose the folder where the images should be exported
-- MUST ALREADY EXISTS
-- set origsFolder to "/Users/stein/Pictures/_tmp_incoming/_apImport"

-- Take rating, color, flag status from Aperture and add as keywords
-- Also, take the face names from the database and add as keyword
set addRatingsAndFacesAsKeywords to false
set ratingsAndColorsParentKeyword to "Organisation"
set facesParentKeyword to "Face"

-- Time to wait after each image is processed (in seconds)
-- Image export takes about 1 second. Waiting time lowers processor load  
-- during long exports.
set delayAfterImage to 0

-- Ignore shell script errors?
set ignoringShellScriptErrors to true

-- Add the version name as IPTC title?
set addVersionNameAsTitle to false

-- Sort pictures into individual camera folders? If false, all images will go into
-- the project folder
set sortIntoCameraFolders to true

-- AppleScript can AFAIK not determine if an image was edited. Use the filter "Adjustments->Has Adjustments" to select 
-- all images with adjustments and manually add the "ApertureEdited" keyword 
set keywordForEditedImages to "ApertureEdited"

-- Add files that Aperture cannot export with embedded IPTC. Aperture will export with XMP sidecar file.
set supportedWithoutIPTC to {"psd", "PSD", "PNG", "png", "gif", "GIF", "pdf", "PDF"}

-- Add files that should not be exported at all here. The script will only
-- run if these are not in the list of selected images
set unsupportedTypes to {"mov", "MOV", "mp4", "MP4", "m4v", "M4V"}

-- Add your Prowl API key here to receive a notification if errors occur or the 
-- export has finished.
set prowlApiKey to ""

-- Set your file naming policy here. The image date is set after the export with exiftool based on the
-- filename. Create a file naming policy in the file export dialog of Aperture that creates files based on 
-- the Aperture picture date in the format YYYYMMD_HH-MM-SS. Add the name of that file naming policy here.
tell application "Aperture"
	set fileNamingPolicy to file naming policy "ImageDateTime"
	copy selection to theSel
end tell

--------------------------- Stop editing here

-- Some definitions to add ratings as keywords 
set colorLabels to {"ColorRed", "ColorOrange", "ColorYellow", "ColorGreen", "ColorBlue", "ColorPurple", "ColorGray"}
set ratings to {"Rating0", "Rating1", "Rating2", "Rating3", "Rating4", "Rating5"}

log "------------------------------- Beginning processing selected images from Aperture"

set numSelected to 0
set numMasters to 0
set numEdited to 0
set numProcessed to 0
set numDuplicated to 0
set numGPS to 0

set startTime to current date



-- Set database names for face name lookups
--if addRatingsAndFacesAsKeywords is equal to true then
set aperture_library_path to do shell script "defaults read com.apple.aperture LibraryPath"
-- expand the '~' if it's in there
set aperture_library_path to do shell script "echo " & aperture_library_path
-- add trailing slash if necessary
if aperture_library_path does not end with "/" then
	set aperture_library_path to aperture_library_path & "/"
end if

-- locate the databases
set the faces_db_path to aperture_library_path & "Database/apdb/faces.db"
set the library_db_path to aperture_library_path & "Database/apdb/Library.apdb"
--end if


-- Pre-flight: check file types
set unsupportedFiles to {}
log "Checking file extensions"
set progress total steps to count of theSel
set progress description to "Checking file types"
set progress completed steps to 0

set numProcessed to 0
repeat with curImg in theSel
	tell application "Aperture"
		set masterFile to (get value of other tag "FileName" of curImg)
		set fileExtension to my suffix(masterFile)
	end tell
	
	if unsupportedTypes contains fileExtension then
		log "Unsupported file: " & masterFile
		set end of unsupportedFiles to masterFile
	end if
	
	-- Add Ratings, Color and Faces as Keywords
	if addRatingsAndFacesAsKeywords is equal to true then
		tell application "Aperture"
			
			-- Set image ratings and labels to keywords		
			set imgRating to (get value of other tag "MainRating" of curImg)
			if imgRating is equal to -1 then
				tell curImg
					make new keyword with properties {name:"Rejected", parents:{ratingsAndColorsParentKeyword}}
				end tell
			else
				tell curImg
					make new keyword with properties {name:(item (imgRating + 1) of ratings), parents:{ratingsAndColorsParentKeyword}}
				end tell
			end if
			
			set imgcolor to (get value of other tag "ColorLabel" of curImg)
			if imgcolor ≤ 0 then
				tell curImg
					make new keyword with properties {name:"ColorNone", parents:{ratingsAndColorsParentKeyword}}
				end tell
			else
				tell curImg
					make new keyword with properties {name:(item imgcolor of colorLabels), parents:{ratingsAndColorsParentKeyword}}
				end tell
			end if
			
			try
				set imgFlagged to (get value of other tag "Flagged" of curImg)
				tell curImg
					make new keyword with properties {name:"Flagged", parents:{ratingsAndColorsParentKeyword}}
				end tell
			on error errStr number errNum
				-- nothing to be done here
			end try
			
			my extract_face_record(curImg, library_db_path, faces_db_path)
		end tell
	end if
	
	
	set numProcessed to numProcessed + 1
	set progress completed steps to numProcessed
end repeat


if (count of unsupportedFiles) is greater than 0 then
	set alertString to my theJoin(unsupportedFiles, return)
	display alert ("Found unsupported files" & return & alertString)
	return
end if

set numSelected to count of theSel
set numProcessed to 0

repeat while (count of theSel) > 0
	
	-- If an image export fails because Aperture still displays a dialog from the previous
	-- image, store it here and process again in the next round
	-- Will lead to duplicates, but if the user dismisses the Aperture window quickly, that
	-- should not be too much	
	set failedExports to {}
	
	-- Every master should be exported only once. Store exported masters here
	set processedMasters to {}
	
	
	set numSel to count of theSel
	
	-- Initialize progress bar
	set progress total steps to numSel
	set progress description to "Exporting"
	set progress completed steps to 0
	
	repeat with curImg in theSel
		
		set remainingTime to my calcRemainingTime(startTime, numProcessed, numSelected + numDuplicated + (count of failedExports))
		set progress additional description to numProcessed & " of " & numSel & " - " & remainingTime
		set numProcessed to numProcessed + 1
		
		delay delayAfterImage
		
		---------------------------------------------------------------------
		-- Get properties from each image
		---------------------------------------------------------------------
		tell application "Aperture"
			get properties of IPTC tags of curImg
			
			set versionName to (get value of other tag "VersionName" of curImg)
			set masterFile to (get value of other tag "FileName" of curImg)
			
			log "========================================================"
			log "Now processing " & numProcessed & " of " & numSel & ": " & masterFile & "(" & my timeStampOf(curImg) & ")"
			-- log "Time remaining: " & remainingTime
			
			get properties of EXIF tags of curImg
			
			set projectName to value of other tag "ProjectName" of curImg
			try
				set cameraMake to value of EXIF tag "Make" of curImg
			on error errStr number errNum
				set cameraMake to "UnknownMake"
			end try
			try
				set cameraModel to value of EXIF tag "Model" of curImg
			on error errStr number errNum
				set cameraModel to "UnknownModel"
			end try
			
			try
				set creatorName to value of IPTC tag "Byline" of curImg
			on error errStr
				set creatorName to ""
			end try
			
		end tell
		
		---------------------------------------------------------------------
		-- Prepare project folders
		---------------------------------------------------------------------
		
		if projectName is "" then
			set projectName to "unknown"
		end if
		
		set projectName to cleanupFolderName(projectName)
		
		set projectYear to makeYearFromProject(projectName)
		
		if projectYear is not equal to "" then
			CheckFolder(origsFolder, projectYear)
			set projectFolder to origsFolder & "/" & projectYear
		else
			set projectFolder to origsFolder
		end if
		
		CheckFolder(projectFolder, projectName)
		
		set projectFolder to projectFolder & "/" & projectName
		
		
		if sortIntoCameraFolders is equal to true then
			set cameraName to cleanupFolderName(cameraMake & " " & cameraModel)
			CheckFolder(projectFolder, cameraName)
			set cameraFolder to projectFolder & "/" & cameraName
			
			my CheckFolder(projectFolder, "_edited")
			set editedFolder to projectFolder & "/_edited"
			
			
		else
			set cameraFolder to projectFolder
			set editedFolder to projectFolder
		end if
		
		
		tell application "Aperture"
			
			tell curImg
				
				set imageKeywords to keywords
				set edited to false
				set exportWithVersionName to false
				
				repeat with k in imageKeywords
					set keywordName to name of k
					if keywordName is keywordForEditedImages then
						set edited to true
					end if
					if keywordName is "ExportWithVersionName" then
						set exportWithVersionName to true
					end if
				end repeat
				
				-- Data to identify the master
				set importGroup to value of other tag "ImportGroup"
				-- set masterId to importGroup & " - " & masterFile
				
				set the master_uuid to every paragraph of (my SQL_command(library_db_path, "select masterUuid from RKVersion where uuid=\"" & id of curImg & "\";"))
				-- log "*** Master IDs Custom: [" & masterId & "] UUID: [" & master_uuid & "]"
				
				set masterId to master_uuid as string
				
				-- GPS information
				set lat to get latitude
				set lon to get longitude
				
			end tell
			
			-- adjust image date imageDate of images {curImg} with masters included
			
			---------------------------------------------------------------------
			-- Export edited and master files
			---------------------------------------------------------------------
			
			if exportWithVersionName is equal to true then
				set fileNamingPolicy to (file naming policy "Version Name")
			else
				set fileNamingPolicy to (file naming policy "ImageDateTime")
			end if
			
			if edited then
				-- Export edited versions
				log "Exporting edited image"
				
				try
					export {curImg} naming files with fileNamingPolicy using export setting "JPEG - Original size" to editedFolder
				on error errMessage
					log "error during export: " & errMessage
					my sendProwlNotification("ApertureExport", "Export error", errMessage, 0)
					set the end of failedExports to curImg
				end try
				
				set numEdited to numEdited + 1
			end if
			
			-- Export original files
			set exportedImages to {}
			if processedMasters does not contain masterId then
				
				set fileExtension to my suffix(masterFile)
				
				if supportedWithoutIPTC contains fileExtension then
					log "Exporting master image without IPTC " & masterId
					-- export {curImg} naming files with fileNamingPolicy to cameraFolder
					
					try
						set exportedImages to (export {curImg} naming files with fileNamingPolicy to cameraFolder metadata sidecar)
					on error errMessage
						log "error during export: " & errMessage
						my sendProwlNotification("ApertureExport", "Export error", errMessage, 0)
						set the end of failedExports to curImg
					end try
					set numMasters to numMasters + 1
				else
					-- JPG or RAW, can be exported with IPTC
					log "Exporting master image with IPTC " & masterId
					
					try
						set exportedImages to (export {curImg} naming files with fileNamingPolicy to cameraFolder metadata embedded)
					on error errMessage
						log "error during export: " & errMessage
						my sendProwlNotification("ApertureExport", "Export error", errMessage, 0)
						set the end of failedExports to curImg
						
					end try
					
					set numMasters to numMasters + 1
				end if
				set end of processedMasters to masterId
			else
				log "Master file does already exist " & masterId
			end if
			
		end tell
		
		---------------------------------------------------------------------
		-- Add GPS and timestamp to the exported master files
		---------------------------------------------------------------------
		get exportedImages
		repeat with exportedImage in exportedImages
			
			if last item of (characters of (exportedImage as string)) is not equal to ":" then
				
				
				log "Updating EXIF data in " & exportedImage
				set imgPath to POSIX path of exportedImage
				
				-- Copy Aperture date from filename to EXIF data
				set exifCommands to "'-datetimeoriginal<filename'"
				-- my doShellScript("/usr/local/bin/exiftool -overwrite_original_in_place \"-datetimeoriginal<filename\" \"" & imgPath & "\"")
				
				-- Add GPS data if present in Aperture
				if lat is not equal to missing value and lon is not equal to missing value then
					set exifCommands to exifCommands & " " & my geotagImage(imgPath, lon, lat)
					set numGPS to numGPS + 1
				end if
				
				-- Add version name as title if so desired
				if addVersionNameAsTitle is equal to true then
					log "Adding version name " & versionName
					set exifCommands to exifCommands & " " & "-title='" & versionName & "'"
					-- my doShellScript("/usr/local/bin/exiftool -overwrite_original_in_place -title=\"" & versionName & "\"  \"" & imgPath & "\"")
				end if
				
				if exifCommands is not equal to "" then
					-- log "/usr/local/bin/exiftool -overwrite_original_in_place " & exifCommands & " '" & imgPath & "'"
					my doShellScript("/usr/local/bin/exiftool -overwrite_original_in_place " & exifCommands & " '" & imgPath & "'")
				end if
			else
				log "Cannot update EXIF of directory " & exportedImage
			end if
		end repeat
		
		set progress completed steps to numProcessed
		
	end repeat
	
	-- run loop again with failed images until everything has been processed.
	-- this will likely result in a few duplicates, but still better than lost information
	
	tell application "Aperture"
		log "Number of failed exports: " & (count of failedExports)
	end tell
	
	set theSel to failedExports
	set numDuplicated to numDuplicated + (count of failedExports)
	
end repeat

set theOut to ("Versions selected: " & numSelected & "
Images processed:	" & numProcessed & "
Masters exported:	" & numMasters & "
Edits exported: " & numEdited & "
Added GPS: " & numGPS & "
Duplicated: " & numDuplicated)

log "All Done!" & return & theOut

display notification theOut with title "Export done"

sendProwlNotification("ApertureExport", "Export finished", theOut, 0)

on CheckFolder(parentFolder, newFolder)
	
	tell application "Finder"
		set f to (parentFolder & "/" & newFolder as POSIX file)
		if exists f then
			-- log "Path " & parentFolder & "/" & newFolder & " already exists."
		else
			log "Creating path " & parentFolder & "/" & newFolder
			make new folder at (POSIX file parentFolder) with properties {name:newFolder}
		end if
	end tell
	
end CheckFolder

-----------------------------------------------------------------------
-- Split a string into a list
on theSplit(theString, theDelimiter)
	-- save delimiters to restore old settings
	set oldDelimiters to AppleScript's text item delimiters
	-- set delimiters to delimiter to be used
	set AppleScript's text item delimiters to theDelimiter
	-- create the array
	set theArray to every text item of theString
	-- restore the old setting
	set AppleScript's text item delimiters to oldDelimiters
	-- return the result
	return theArray
end theSplit

-----------------------------------------------------------------------
-- Join a list to a single string using given delimiter
on theJoin(theList, theDelimiter)
	-- save delimiters to restore old settings
	set oldDelimiters to AppleScript's text item delimiters
	-- set delimiters to delimiter to be used
	set AppleScript's text item delimiters to theDelimiter
	-- create the array
	set theString to theList as string
	-- restore the old setting
	set AppleScript's text item delimiters to oldDelimiters
	-- return the result
	return theString
end theJoin

on suffix(theString)
	
	set theReversedFileName to (reverse of (characters of theString)) as string
	set theOffset to offset of "." in theReversedFileName
	set thePrefix to (reverse of (characters (theOffset + 1) thru -1 of theReversedFileName)) as string
	set theSuffix to (reverse of (characters 1 thru (theOffset - 1) of theReversedFileName)) as string
	
	return theSuffix
	
end suffix


on extract_face_record(this_photo, library_db_path, faces_db_path)
	set the master_uuid to every paragraph of (my SQL_command(library_db_path, "select masterUuid from RKVersion where uuid=\"" & id of this_photo & "\";"))
	--get detected faces for the image key
	set the face_keys to every paragraph of (my SQL_command(faces_db_path, "select faceKey from RKDetectedFace where masteruuid=\"" & master_uuid & "\";"))
	
	-- for each face
	set the face_records to {}
	repeat with this_key in the face_keys
		-- get name. NOTE could select fullName instead of select name if you wished
		set the face_name to my SQL_command(faces_db_path, "select name from rkFaceName where faceKey=\"" & this_key & "\";")
		-- add it as new key word
		
		if face_name is not equal to "" then
			tell application "Aperture"
				tell this_photo
					make new keyword with properties {name:face_name, parents:{"Face"}}
				end tell
			end tell
		else
			-- log "No face in this image"
		end if
		
	end repeat
end extract_face_record

-----------------------------------------------------------------------
-- Try repeatedly to get a value from a database. Can fail if Aperture has the dabase locked,
-- so we just keep trying. Failure is not an option ;-)
on SQL_command(faces_db_path, command_string)
	
	set theResult to "---NOFACE---"
	repeat while theResult is equal to "---NOFACE---"
		try
			set theResult to (do shell script "sqlite3 " & (quoted form of faces_db_path) & " '" & command_string & "'")
		on error errStr number errNum
			log "Database " & faces_db_path & " is locked. Waiting before trying again..."
			delay 1
		end try
	end repeat
	
	return theResult
	
end SQL_command

-----------------------------------------------------------------------
-- Execute a shell comman, catch the output and log
on doShellScript(theCommand)
	
	if my ignoringShellScriptErrors = true then
		try
			do shell script theCommand
		on error errMessage
			log "Warning: " & errMessage
		end try
	else
		do shell script theCommand
	end if
	
end doShellScript

-----------------------------------------------------------------------
-- Convert decimal comma to dot
on commaToDecimalPoint(theNumber)
	
	set theString to theNumber as string
	
	set oldDelimiters to AppleScript's text item delimiters
	set AppleScript's text item delimiters to ","
	set theList to text items of theString
	
	set AppleScript's text item delimiters to "."
	set theResult to theList as string
	set AppleScript's text item delimiters to oldDelimiters
	
	return theResult
	
end commaToDecimalPoint

-----------------------------------------------------------------------
-- Create a string for exiftool to add GPS info to image
on geotagImage(imgPath, theLongitude, theLatitude)
	
	if theLatitude < 0 then
		set latRef to "S"
		set lat to -theLatitude
	else
		set latRef to "N"
		set lat to theLatitude
	end if
	
	if theLongitude < 0 then
		set lonRef to "W"
		set lon to -theLongitude
	else
		set lonRef to "E"
		set lon to theLongitude
	end if
	
	set latStr to my commaToDecimalPoint(lat)
	set lonStr to my commaToDecimalPoint(lon)
	
	-- log "Image location: " & latStr & latRef & ", " & lonStr & lonRef
	
	-- log "/usr/local/bin/exiftool -overwrite_original_in_place -P -exif:GPSLatitude=" & latStr & " -exif:GPSLatitudeRef=" & latRef & " -exif:GPSLongitude=" & lonStr & " -exif:GPSLongitudeRef=" & lonRef & "  " & imgPath
	-- my doShellScript("/usr/local/bin/exiftool -overwrite_original_in_place -P -exif:GPSLatitude=" & latStr & " -exif:GPSLatitudeRef=" & latRef & " -exif:GPSLongitude=" & lonStr & " -exif:GPSLongitudeRef=" & lonRef & " " & "\"" & imgPath & "\"")
	
	set exifString to "-exif:GPSLatitude=" & latStr & " -exif:GPSLatitudeRef=" & latRef & " -exif:GPSLongitude=" & lonStr & " -exif:GPSLongitudeRef=" & lonRef
	return exifString
	
end geotagImage

-----------------------------------------------------------------------
-- Calculate time based on time passed since start of script
on calcRemainingTime(startTime, itemsDone, numItems)
	
	if itemsDone is less than 1 then
		return "unknown"
	end if
	
	set timePerItem to ((current date) - startTime) / itemsDone
	set remainingTime to timePerItem * (numItems - itemsDone)
	
	set remainingTime to (((remainingTime * 100) + 0.5) div 1) / 100
	set remainingTime to round (remainingTime)
	set remainingTimeStr to formatAsTime(remainingTime)
	
	--set finishTime to (get short date string of ((current date) + remainingTime)) & " " & (get time string of ((current date) + remainingTime))
	set finishTime to (get time string of ((current date) + remainingTime))
	return remainingTimeStr & " until " & finishTime
	
end calcRemainingTime

-----------------------------------------------------------------------
-- Send notificating to my iPhone
on sendProwlNotification(theApp, theTitle, theText, intPriority)
	if my prowlApiKey is not equal to "" then
		try
			do shell script "curl -s -k https://api.prowlapp.com/publicapi/add -F apikey=" & prowlApiKey & " -F application='" & theApp & "' -F event='" & theTitle & "' -F description='" & theText & "' -F priority=" & intPriority
		on error errMsg
		end try
	end if
end sendProwlNotification

-----------------------------------------------------------------------
-- Format number of seconds to H:MM:SS
on formatAsTime(theSeconds)
	
	set resHours to round (theSeconds div hours)
	set resSeconds to (theSeconds mod hours)
	set resMinutes to round (resSeconds div minutes)
	set resSeconds to (resSeconds mod minutes)
	
	if resMinutes < 10 then
		set resMinutes to "0" & (round (resMinutes) as string)
	end if
	
	if resSeconds < 10 then
		set resSeconds to "0" & resSeconds as string
	end if
	
	return (resHours & ":" & resMinutes & ":" & resSeconds as string)
end formatAsTime

-- ---------------------------------------------------------------------
-- Get the first 4 characters of a string and check if they are
-- digits. If so, return, otherwise return empty string
on makeYearFromProject(theProject)
	
	try
		set prefix to characters 1 thru 4 of theProject as string
		set projYear to prefix as number
		return prefix
	on error errMsg
		-- log errMsg
		return ""
	end try
	
end makeYearFromProject

on cleanupFolderName(folderName)
	set allowedCharacters to "ABCDEFGHIJKLMNOPQRSTUVWXYZ abcdefghijklmnopqrstuvwxyz0123456789-_,.äöüÄÖÜß"
	
	set the new_text to ""
	repeat with this_char in folderName
		set x to the offset of this_char in the allowedCharacters
		if x is not 0 then
			set the new_text to (the new_text & this_char) as string
		end if
	end repeat
	return the new_text
end cleanupFolderName

on timeStampOf(theImage)
	
	tell application "Aperture"
		tell theImage
			try
				set theYear to value of EXIF tag "CaptureYear"
			on error errString
				set theYear to "0000"
			end try
			try
				set theMonth to value of EXIF tag "CaptureMonthOfYear"
			on error errString
				set theMonth to "00"
			end try
			try
				set theDay to value of EXIF tag "CaptureDayOfMonth"
			on error errString
				set theDay to "00"
			end try
			try
				set theHour to value of EXIF tag "CaptureHourOfDay"
			on error errString
				set theHour to "00"
			end try
			try
				set theMinute to value of EXIF tag "CaptureMinuteOfHour"
			on error errString
				set theMinute to "00"
			end try
			try
				set theSecond to value of EXIF tag "CaptureSecondOfMinute"
			on error errString
				set theSecond to "00"
			end try
		end tell
	end tell
	
	set theDate to (my fillLeft(theYear, "0", 4))
	set theDate to theDate & "/" & (my fillLeft(theMonth, "0", 2))
	set theDate to theDate & "/" & (my fillLeft(theDay, "0", 2))
	set theDate to theDate & " " & (my fillLeft(theHour, "0", 2))
	set theDate to theDate & ":" & (my fillLeft(theMinute, "0", 2))
	set theDate to theDate & ":" & (my fillLeft(theSecond, "0", 2))
	
end timeStampOf

on fillLeft(theNumber, theFillChar, theLength)
	
	set theString to (theNumber as string)
	set i to offset of "." in theString
	if i > 0 then
		set theString to (characters 1 through (i - 1) of theString)
	end if
	set i to offset of "," in theString
	if i > 0 then
		set theString to (characters 1 through (i - 1) of theString)
	end if
	repeat while length of theString is less than theLength
		set theString to theFillChar & theString
	end repeat
	
	return theString
	
end fillLeft
