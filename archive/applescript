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
	
	if revNameNoExt is "" or revNameNoExt starts with "." then
		set revNameNoExt to "comparison"
	end if
	
	set outName to revNameNoExt & ".redline.docx"
	set tempOutPOSIX to desktopPOSIX & outName
	set finalOutPOSIX to directoryPath & outName
	
	-- Copy files to Desktop
	do shell script "cp " & quoted form of (POSIX path of origAlias) & " " & quoted form of tempOrigPOSIX
	do shell script "cp " & quoted form of revPOSIX & " " & quoted form of tempRevPOSIX
	
	-- Wait for files to be fully written
	delay 2
	
	-- Verify files exist and are readable
	do shell script "ls -la " & quoted form of tempOrigPOSIX
	do shell script "ls -la " & quoted form of tempRevPOSIX
	
	-- Quit Word completely first, then restart
	tell application "Microsoft Word"
		quit
	end tell
	
	delay 2
	
	tell application "Microsoft Word"
		activate
		delay 1
		
		-- Open both documents to accept all tracked changes
		set origDoc to open file name tempOrigPOSIX
		delay 1
		set revDoc to open file name tempRevPOSIX
		delay 1
		
		-- Accept all tracked changes in both documents
		try
			tell origDoc
				accept all revisions
			end tell
		end try
		
		try
			tell revDoc
				accept all revisions
			end tell
		end try
		
		delay 1
		
		-- Save both documents with changes accepted
		save origDoc
		save revDoc
		
		-- Close the revised document (compare will reopen it)
		close revDoc saving no
		
		delay 1
		
		-- Try the compare
		try
			compare origDoc path tempRevPOSIX
		on error errMsg
			display dialog "Compare failed: " & errMsg & return & return & "Files: " & return & origName & return & revName
			close every document saving no
			do shell script "rm " & quoted form of tempOrigPOSIX
			do shell script "rm " & quoted form of tempRevPOSIX
			return input
		end try
		
		delay 5
		
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
	
	-- Use Finder to move file to target location
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
	
	if fileName is "" then
		set AppleScript's text item delimiters to oldDelims
		return ""
	end if
	
	if fileName does not contain "." then
		set AppleScript's text item delimiters to oldDelims
		return fileName
	end if
	
	set AppleScript's text item delimiters to "."
	set nameItems to text items of fileName
	set itemCount to count of nameItems
	
	if itemCount is 2 and item 1 of nameItems is "" then
		set AppleScript's text item delimiters to oldDelims
		return fileName
	end if
	
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
