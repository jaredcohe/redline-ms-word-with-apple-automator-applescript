on run {input, parameters}
	-- require exactly two files
	log "=== Starting Word Compare ==="
	
	if (count of input) is not 2 then
		display dialog "Select exactly two Word files." buttons {"OK"} default button 1
		return input
	end if
	
	set origAlias to item 1 of input as alias
	set revAlias to item 2 of input as alias
	
	tell application "System Events"
		set origName to name of origAlias
		set revName to name of revAlias
	end tell
	
	-- output path next to the revised file
	set revPOSIX to POSIX path of revAlias
	set outPOSIX to revPOSIX & ".redline.docx"
	
	tell application "Microsoft Word"
		activate
		
		-- open original document
		set origDoc to open file name (POSIX path of origAlias)
		
		-- Compare: creates a new document with tracked changes
		-- The new document becomes the active document
		compare origDoc path (POSIX path of revAlias)
		
		-- Get the newly created comparison document
		set compDoc to active document
		
		-- optional: stamp author for the revisions
		try
			set user name to "Auto Redline"
		end try
		
		-- save the comparison result
		save as compDoc file name outPOSIX file format format document default
		
		-- close the original document
		close origDoc saving no
		
		-- optionally close the comparison document
		-- close compDoc saving no
		
	end tell
	
	return outPOSIX
end run
