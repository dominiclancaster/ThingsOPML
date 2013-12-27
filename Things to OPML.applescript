# Export Things to OPML - for exporting the Things database as an OPML file
# Dexter Ang - @thepoch on Twitter
# Copyright (c) 2013, Dexter Ang
# 
# Tested with Things 2.2.5 on OS X Mavericks 10.9.1
# 

-- ---------- Notes ----------
-- "This is a note !@#$%^&*(){}:\"<>?" --- this string doesn't escape properly using python. Why?
-- --------------------------------------------------

-- ---------- Variables ----------
property shouldLogCompleted : false
property shouldEmptyTrash : false
set sourceLists to {"Inbox", "Today", "Next", "Scheduled", "Someday", "Projects"}
property notificationEnabled : true
property notifyOnStart : true
property notifyOnEnd : true
-- --------------------------------------------------

-- ---------- Constants ----------
set opmlStart to "<?xml version=\"1.0\" encoding=\"UTF-8\"?><opml version=\"1.0\"><head><title></title></head><body>"
set outlineString to "<outline text=%s _status=%s _note=%s"
set outlineWithoutSub to "/>"
set outlineWithSub to ">"
set outlineStringClose to "</outline>"
set opmlEnd to "</body></opml>"
-- --------------------------------------------------

-- ---------- The actual algorithm ----------
set opmlXML to opmlStart

tell application "Things"
	-- Save file to --
	set plainTextFile to choose file name with prompt "Select location and filename. Extension will always be .opml." default name "Things.opml"
	set plainTextFile to plainTextFile as text
	if plainTextFile does not end with ".opml" then
		set plainTextFile to plainTextFile & ".opml"
	end if
	set outputFile to POSIX path of plainTextFile
	if outputFile is "" then return
	try
		open for access outputFile with write permission
		set eof outputFile to 0
	on error
		close access outputFile
		display dialog "Error: cannot write to file."
		return
	end try
	
	if shouldLogCompleted then log completed now
	if shouldEmptyTrash then empty trash
	
	if notificationEnabled and notifyOnStart then
		notifyme("Exported has started") of me
	end if
	
	repeat with aSource in sourceLists
		set content to formatted_string(outlineString, {quoted form of (aSource as string), quoted form of "", quoted form of ""}) of me & outlineWithSub
		
		set taskItems to to dos of list aSource
		repeat with taskItem in taskItems
			set rawTaskTitle to the name of taskItem -- will be used for "Projects" below".
			set taskTitle to urlencode(the name of taskItem) of me
			
			--			display dialog "rawTaskTitle: " & rawTaskTitle
			--			display dialog "taskTitle: " & taskTitle
			
			--			display dialog "escapedoublequotes: " & escapedoublequotes(the name of taskItem) of me
			--			display dialog "urlencode and escapedoublequotes: " & urlencode(escapedoublequotes(the name of taskItem) of me) of me
			--			display dialog "remotequotes and urlencode and escapedoublequotes: " & removequotes(urlencode(escapedoublequotes(the name of taskItem) of me) of me) of me
			
			--			if due date of taskItem is not missing value then set taskTitle to taskTitle & " (Due: " & date string of (due date of taskItem) & ")"
			set taskChecked to quoted form of ""
			if status of taskItem is completed then set taskChecked to quoted form of "checked"
			set rawTaskNotes to the notes of taskItem
			if notes of taskItem is not "" then
				set taskNotes to urlencode(the notes of taskItem) of me
			else
				set taskNotes to quoted form of ""
			end if
			--			display dialog rawTaskTitle & " " & taskChecked & " " & the quoted form of rawTaskNotes
			--			display dialog formatted_string2(outlineString, {rawTaskTitle, taskChecked, rawTaskNotes}) of me
			set content to content & linefeed & formatted_string(outlineString, {taskTitle, taskChecked, taskNotes}) of me
			
			if (aSource as string = "Projects") then
				set content to content & outlineWithSub
			else
				set content to content & outlineWithoutSub
			end if
			
			-- Special case for Projects.
			if (aSource as string = "Projects") then
				set projectTaskItems to to dos of project rawTaskTitle
				repeat with projectTaskItem in projectTaskItems
					set projectTaskTitle to urlencode(the name of projectTaskItem) of me
					--					if due date of projectTaskItem is not missing value then set projectTaskTitle to projectTaskTitle & " (Due: " & date string of (due date of projectTaskItem) & ")"
					set projectTaskChecked to quoted form of ""
					if status of projectTaskItem is completed then set projectTaskChecked to quoted form of "checked"
					if notes of projectTaskItem is not "" then
						set projectTaskNotes to urlencode(the notes of projectTaskItem) of me
					else
						set projectTaskNotes to quoted form of ""
					end if
					set content to content & linefeed & formatted_string(outlineString, {projectTaskTitle, projectTaskChecked, projectTaskNotes}) of me & outlineWithoutSub
				end repeat
				set content to content & linefeed & outlineStringClose
			end if
		end repeat
		
		set content to content & linefeed & outlineStringClose
		
		set opmlXML to opmlXML & linefeed & content
	end repeat
	
	set opmlXML to opmlXML & linefeed & opmlEnd
	
	write {opmlXML} as Çclass utf8È to outputFile
	close access outputFile
	
	if notificationEnabled and notifyOnEnd then
		notifyme("Exported as ended") of me
	end if
end tell
-- --------------------------------------------------

-- ---------- Helper functions ----------
-- From: http://www.tow.com/2006/10/12/applescript-stringwithformat/
on formatted_string(key_string, parameters)
	set cmd to "printf " & quoted form of key_string
	repeat with i from 1 to count parameters
		set cmd to cmd & space & quoted form of ((item i of parameters) as string)
	end repeat
	return do shell script cmd
end formatted_string

on formatted_string2(key_string, parameters)
	set cmd to "env python -c 'import sys; from xml.sax.saxutils import quoteattr; print " & key_string & " % ("
	repeat with i from 1 to count parameters
		set cmd to cmd & space & "quoteattr(" & ((item i of parameters) as string) & "),"
	end repeat
	set cmd to cmd & ")'"
	return do shell script cmd
end formatted_string2

-- From: https://discussions.apple.com/message/9380471#9380471
on urlencode(parameter)
	set cmd to "/usr/bin/python -c 'import sys; from xml.sax.saxutils import quoteattr; print quoteattr(sys.argv[1])' " & quoted form of parameter
	return do shell script cmd
end urlencode

-- This function removes the extra quotes of what is returned by the urlencode function above.
on removequotes(parameter)
	return text 2 thru ((length of parameter) - 1) of parameter
end removequotes

on escapedoublequotes(parameter)
	return do shell script "echo " & quoted form of parameter & " | sed -e 's/\"/\\\\\"/g'"
end escapedoublequotes

on notifyme(parameter)
	-- code to check for OS version. UNUSED.
	set os_version to do shell script "sw_vers -productVersion"
	
	display notification parameter as text with title "Export Things to OPML"
end notifyme
-- --------------------------------------------------
