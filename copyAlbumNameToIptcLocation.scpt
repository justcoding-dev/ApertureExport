-- Take the project of the first image in the selection, get the project for image
-- Find folder "Locations" in that project, then go through all albums in this folder.
-- For each image in every album, set the name of the album as "Image Location" IPTC tag
-- (which is "SubLocation" in Aperture")
--
-- Example?
--
-- ProjectX
-- |- Folder "Locations"
--       |-- Album "Beach" 
--       |-- Album "Museum"
--
-- When selecting any picture in the project, the script will apply the keyword "Beach" to 
-- all images in the "Beach" folder and "Museum" to all images in the "Museum" folder.
--

tell application "Aperture"
	
	copy selection to theSelection
	
	set theImage to first item of theSelection
	
	set theProject to get value of other tag "MasterProject" of theImage
	
	set datedFolder to subfolder "Locations" of theProject
	set dateAlbums to every album of datedFolder
end tell

set progress total steps to (count of dateAlbums)
set progress description to "Adding album name to IPTC location"
set progress completed steps to 0
set numProcessed to 0

repeat with theAlbum in dateAlbums
	
	
	set albumName to name of theAlbum
	
	log "Processing images in album " & albumName
	tell application "Aperture"
		set theVersions to every image version in theAlbum
		
		repeat with theVersion in theVersions
			
			log (name of theVersion) & ": " & albumName
			
			tell theVersion
				make new IPTC tag with properties {name:"SubLocation", value:albumName}
			end tell
		end repeat
	end tell
	
	set numProcessed to numProcessed + 1
	set progress completed steps to numProcessed
	
	
end repeat

