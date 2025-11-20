on run {input, parameters}
	-- require exactly two files
	log "=== Starting Word Compare ==="
	
	if (count of input) is not 2 then
		display dialog "Select exactly two Word files." buttons {"OK"} default button 1
		return input
	end if
	
	set origAlias to item 1 of input as alias
	set revAlias to item 2 of input as alias
	
	-- Get filenames
	tell application "System Events"
		set origName to name of origAlias
		set revName to name of revAlias
	end tell
	
	log "Original file: " & origName
	log "Revised file: " & revName
	
	-- Get the revised file's directory for final output
	set revPOSIX to POSIX path of revAlias
	set oldDelims to AppleScript's text item delimiters
	set AppleScript's text item delimiters to "/"
	set pathParts to text items of revPOSIX
	set directoryParts to items 1 thru -2 of pathParts
	set directoryPath to (directoryParts as text) & "/"
	set AppleScript's text item delimiters to oldDelims
	
	-- Build temp paths on Desktop
	set desktopPOSIX to POSIX path of (path to desktop)
	set tempOrigPOSIX to desktopPOSIX & origName
	set tempRevPOSIX to desktopPOSIX & revName
	
	-- Strip extension properly and validate
	set revNameNoExt to my stripExtension(revName)
	log "Filename without extension: '" & revNameNoExt & "'"
	
	-- Safety check: if empty or starts with dot, use original name
	if revNameNoExt is "" or revNameNoExt starts with "." then
		set revNameNoExt to "comparison"
		log "WARNING: Invalid filename, using 'comparison' instead"
	end if
	
	set outName to revNameNoExt & ".redline.docx"
	log "Output filename: " & outName
	
	set tempOutPOSIX to desktopPOSIX & outName
	set finalOutPOSIX to directoryPath & outName
	
	-- Copy files to Desktop
	do shell script "cp " & quoted form of (POSIX path of origAlias) & " " & quoted form of tempOrigPOSIX
	do shell script "cp " & quoted form of revPOSIX & " " & quoted form of tempRevPOSIX
	
	tell application "Microsoft Word"
		activate
		
		-- Open documents from Desktop
		set origDoc to open file name tempOrigPOSIX
		
		-- Compare
		compare origDoc path tempRevPOSIX
		
		delay 2
		
		-- Get the comparison document
		set compDoc to active document
		
		-- Save to Desktop
		save as compDoc file name tempOutPOSIX file format format document
		
		-- Close documents
		close origDoc saving no
		close compDoc saving no
		
	end tell
	
	-- Clean up temp input files
	do shell script "rm " & quoted form of tempOrigPOSIX
	do shell script "rm " & quoted form of tempRevPOSIX
	
	-- Use Finder to move file to Google Drive
	tell application "Finder"
		set sourceFile to POSIX file tempOutPOSIX as alias
		set targetFolder to POSIX file directoryPath as alias
		move sourceFile to targetFolder with replacing
	end tell
	
	return finalOutPOSIX
end run

-- Helper to strip file extension
on stripExtension(fileName)
	set oldDelims to AppleScript's text item delimiters
	
	-- Handle edge cases
	if fileName is "" then
		set AppleScript's text item delimiters to oldDelims
		return ""
	end if
	
	-- If no dot, return as-is
	if fileName does not contain "." then
		set AppleScript's text item delimiters to oldDelims
		return fileName
	end if
	
	-- Split on dots
	set AppleScript's text item delimiters to "."
	set nameItems to text items of fileName
	set itemCount to count of nameItems
	
	-- If filename starts with dot (like .hidden), return as-is
	if itemCount is 2 and item 1 of nameItems is "" then
		set AppleScript's text item delimiters to oldDelims
		return fileName
	end if
	
	-- Take all items except the last one (the extension)
	if itemCount > 1 then
		set resultItems to items 1 thru (itemCount - 1) of nameItems
		set AppleScript's text item delimiters to "."
		set resultText to resultItems as text
		set AppleScript's text item delimiters to oldDelims
		return resultText
	else
		set AppleScript's text item delimiters to oldDelims
		return fileName
	end if
end stripExtension
