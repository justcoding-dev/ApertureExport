-- Take the project of the first image in the selection, get the project for image
-- Find all folders that begin with "keyword-" in that project, then go through all albums in this folder.
-- For each image in every album, set the name of the album as keyword with the words in the 
-- parent folder as parent keywords
-- (which is "SubLocation" in Aperture")
--
-- Example?
--
-- ProjectX
-- |- Folder "keyword-Actions"
--       |-- Album "Swimming" 
--       |-- Album "Dancing"
--
-- When selecting any picture in the project, the script will apply the keyword "Actions->Swimming" to 
-- all images in the "Swimming" folder and "Actions->Dancing" to all images in the "Dancing" folder.

tell application "Aperture"
	
	copy selection to theSelection
	
	set theImage to first item of theSelection
	
	set theProject to get value of other tag "MasterProject" of theImage
	
	set keywordFolders to every subfolder of theProject whose name begins with "Keyword-"
end tell

set progress total steps to (count of keywordFolders)
set progress description to "Adding keywords to images"
set progress completed steps to 0
set numProcessed to 0

repeat with keywordFolder in keywordFolders
	
	tell application "Aperture"
		
		set theAlbums to every album of keywordFolder
		get properties of keywordFolder
		set foldername to (name of keywordFolder)
		set parentKeywords to my getParentKeywords(foldername)
		
		repeat with theAlbum in theAlbums
			
			
			set albumName to name of theAlbum
			
			log "Processing images in album " & albumName
			set theVersions to every image version in theAlbum
			
			repeat with theVersion in theVersions
				
				log (name of theVersion) & ": " & albumName & ", " & parentKeywords
				
				tell theVersion
					
					make new keyword with properties {name:albumName, parents:parentKeywords}
					
				end tell
			end repeat
			
			
		end repeat
		
	end tell
	
	set numProcessed to numProcessed + 1
	set progress completed steps to numProcessed
	
end repeat

on getParentKeywords(foldername)
	
	set theList to my theSplit(foldername, "-")
	if (count of theList) > 1 then
		return theJoin(reverse of (items 2 through (count of theList) of theList), "	")
	end if
	
	return {}
end getParentKeywords

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

