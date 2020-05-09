(*
 * This program is free software. It comes without any warranty,
 * to the extent permitted by applicable law. You can redistribute it
 * and/or modify it under the terms of the Don't Be A Dick Public License,
 * version 1.1, as published by Phil Sturgeon. See this project's
 * COPYING.txt file or http://dbad-license.org for more details.
 *)




------------------------------------------------------
------------------------------------------------------
------ Set global constants and import user settings
------ These constants should not need updating
------  unless there are changes to the HTML
------    of MyAccess or WMT Scheduler
------------------------------------------------------
------------------------------------------------------

-- how long to wait, in seconds, before starting to check if a webpage has loaded
-- edit this property if you consistently have problems with the script running before a page loads
property globalPreDelay : 0

-- constants related to logging in to MyAccess
property webScheduleURL : "https://wmtscheduler.faa.gov/WMT_LogOn/"
property wmtLoginButtonID : "btnLogin"
property emailInputID : "userEmail"
property emailButtonID : "cont"
property pinInputID : "PIN_INPUT"
property answerInputID : "answer0"
property myAccessLoginButtonID : "LOGIN_BUTTON"

-- strings used to check which page is loaded
property strInWMTSplashPage : "WARNING!!! WARNING!!! WARNING!!!"
property strInEmailEntryPage : "Use Your Email Address"
property strInPinEntryPage : "MyAccess PIN"
property strInWorksheetView : "<title>Worksheet View</title>"
property strInScheduleView : "<title>My Schedule</title>"
property strInLegendView : "<title>Shift Legend</title>"

-- constants for navigating within the WMT Scheduler website
property scheduleViewURL : "Index.asp?Action=actmySchedule"
property shiftLegendURL : "Index.asp?Action=dspshiftLegend"

-- constants for getting the shift legends
-- there are six centered header cells in the main table
-- ...and six cells per row, but the "Remarks" cell is not centered
property numCenteredHeadingCells : 6
property numCenteredContentCells : 5

-- constants for getting shifts from "my schedule" view
-- there are five childNodes in each cell's div: weekday, <br>, date, <br>, shift
-- ...we want the date and the shift, zero-indexed for JavaScript
property indexOfDateInScheduleDiv : 2
property indexOfShiftInScheduleDiv : 4
property daysInPP : 14

-- constants regarding shift ID format
property digits : {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}
property oneCharPrepends : {"C", "D", "T"}
property fourCharPrepends : {"AREA"}
property oneCharAppends : {"$", "^", "!"}
property twoCharAppends : {"OM"}
property threeCharAppends : {"TMC"}
property adminChar : "@"
property offsiteChar : "<"

-- constants for GUI text
property loginPromptTitle : "WSTC: MyAccess Login Settings"
property settingsPromptTitle : "WSTC: Script Settings"
property welcomeText : "Welcome to the WebScheduleToCalendar script! The script will do the following:

1. Ask your permission to use various applications and settings. More information can be found in the README file.
2. Ask for your desired configuration settings.
3. Ask for your MyAccess login information.
4. Get your current- and upcoming-pay-period shifts and enter them into your calendar.

The login information will be encrypted and stored locally for use next time you run the script. It is not used for any other purpose."
property titlePrompt : "Enter the event title to be used for the work events:"
property calPrompt : "Choose which calendar to use for the work events:"
property addressPrompt : "Enter your facility's address if desired, or leave blank:"
property appendPrompt : "Should the event titles include the shift ID?"
property populatePrompt : "Should the event descriptions include the shift ID?"
property closeSafariPrompt : "The script will open the WMT Scheduler website.
Should it close the website when it is done?"
property closeCalendarPrompt : "The script will open Calendar.
Should it close Calendar when it is done?"
property emailPrompt : "Enter the email address you use for MyAccess:"
property pinPrompt : "Enter the PIN you use for MyAccess:"
property questionPrompt : "Enter your answer to the question: 

"
property badPassPrompt : "The script has been recompiled since the last time it was run.
The key used to encrypt your login information is no longer valid.

Please re-enter the PIN you use for MyAccess:"
property badEmailPrompt : "MyAccess has rejected your email address.

Please re-enter the email address you use for MyAccess:"
property badPinPrompt : "MyAccess has rejected the PIN and/or secret answer.

Please re-enter the PIN you use for MyAccess:"
property strAllowJSinAE : "The script needs to run JavaScript in Safari in order to apply your login information and get the shifts from WMT Scheduler. The system will ask for an administrator's username and password to allow this.

The script can click through the menu items for you, which will require permission to control accessibility elements, or you can do the process manually by following the linked instructions. Which would you like?"
property strWaitForJS : "Please click Continue once you have allowed JavaScript access in Safari"
property strWaitForUser : "Waiting for user input…"
property strScriptError : "The script encountered an error and must quit now. Several common errors boil down to the script running too quickly and/or a page being slow to load; try running the script one or two more times to make sure the error is persistent.

The error was:
"

-- URL to show the user how to enable JavaScript from Apple Events
property allowJSURL : "https://github.com/AnonymousController/wstc/blob/master/AllowJS.md"

-- variables that the main handlers will share among themselves
property scriptName : "WebScheduleToCalendar"
global settingsPath, databasePath, legendPath
global s_plist, l_plist, db_plist
global eventRec, shiftRec
global eventTitle, calName, address, appendShift, populateDescription
global closeSafariWhenDone, closeCalendarWhenDone
global email, encPin, safariWasRunning, calendarWasRunning
global currentStartDate

-- pw is re-initialized every time the script compiles
property pw : do shell script "openssl rand -base64 30"




------------------------------------------------------
------------------------------------------------------




------------------------------------------------------
------------------------------------------------------
------ HERE IS THE MAIN FUNCTION
------------------------------------------------------
------------------------------------------------------
on run {}
	
	try
		-- set up "progress" display
		set progress total steps to 55
		set progress completed steps to 0
		set progress description to "Setting up…"
		
		-- set things up for our script
		-- this includes first-run stuff (access, app support dir)
		-- and each-run stuff (application states, plist paths)
		initialize()
		set progress completed steps to 11
		set progress additional description to ""
		set progress description to "Getting user settings"
		
		-- get settings for the script from user, or load from disk
		if not (exists file settingsPath of application "Finder") then
			create_settings()
		else
			set progress completed steps to 16
			set progress description to "Loading user settings"
			import_settings()
		end if
		set progress completed steps to 21
		set progress description to "Logging in to MyAccess…"
		
		-- log in to WMT Scheduler via MyAccess
		login()
		set progress completed steps to 31
		set progress additional description to ""
		set progress description to "Generating list of shifts from WMT Scheduler"
		
		-- scrape all shifts from current and future PPs from WMT Scheduler
		get_shifts()
		set progress completed steps to 41
		set progress additional description to ""
		set progress description to "Creating calendar events"
		
		-- create calendar events based on the shifts
		create_events()
		set progress completed steps to 51
		set progress description to "Finishing"
		
		-- clean up database file
		clean_up_db()
		
		-- close applications
		close_apps()
		
		set progress completed steps to 55
		
	on error errStr number errNum
		display dialog strScriptError & errNum & ": " & errStr buttons {"Quit"} default button 1 with title "WSTC"
	end try
end run




------------------------------------------------------
------------------------------------------------------




------------------------------------------------------
------------------------------------------------------
------ Setup: permissions, file paths, etc
------------------------------------------------------
------------------------------------------------------
to initialize()
	set safariWasRunning to (application "Safari" is running)
	set calendarWasRunning to (application "Calendar" is running)
	
	set appSupportPath to (path to application support from user domain as text)
	set myAppSupportPath to appSupportPath & scriptName & ":"
	
	set settingsPath to myAppSupportPath & "Settings.plist"
	set databasePath to myAppSupportPath & "Database.plist"
	set legendPath to myAppSupportPath & "ShiftLengends.plist"
	
	set progress completed steps to 1
	
	-- check if the initialization steps have been run before
	set initialized to false
	tell application "Finder" to if (exists folder myAppSupportPath) then ¬
		set initialized to true
	
	if not initialized then
		-- welcome user
		display dialog welcomeText with title scriptName buttons {"OK"} default button 1 with icon note
		
		-- create our Application Support folder if necessary
		-- this triggers the Finder Access request
		set progress additional description to "Requesting Finder access…"
		delay 0.05
		tell application "Finder" to make new folder at appSupportPath with properties {name:scriptName}
		set progress completed steps to 2
		
		-- trigger System Events request
		set progress additional description to "Requesting System Events access…"
		delay 0.05
		tell application "System Events" to tell first process to set foo to name
		set progress completed steps to 3
		
		-- access calendars to trigger calendar request
		set progress additional description to "Requesting Calendar access…"
		delay 0.05
		if not (calendarWasRunning) then my_activate_return("Calendar", false)
		tell application "Calendar" to tell first calendar to set foo to name
		set progress completed steps to 4
		
		-- attempt to give ourselves permission to use JS in Safari
		-- unfortunately ever since macOS 10.11 or so this won't work
		-- because the user needs to enter an admin password to enable JS from AE
		-- and there doesn't seem to be a way to trigger that besides the menubar GUI:
		-- the value stored in the plist will change, but the functionality will not be there.
		-- note that if the user has previously turned the setting on, but now it is off,
		-- this should work to turn it back on.
		set progress additional description to "Requesting Safari access…"
		delay 0.05
		do shell script "defaults write -app Safari AllowJavaScriptFromAppleEvents -bool true"
		
		-- now try to actually use JavaScript in Safari to see if the setting worked
		try
			if not (safariWasRunning) then my_activate_return("Safari", false)
			set progress completed steps to 5
			-- NOTE: we don't actually need a document to be open in order to test
			-- and telling it to tell document 1 throws an error if no window is open
			tell application "Safari" to do JavaScript "var foo = 0;"
			
		on error -- nope, setting didn't work, and we got an error saying so!
			-- ask user how they want to proceed
			set progress additional description to "Requesting JavaScript access in Safari…"
			delay 0.05
			display dialog strAllowJSinAE buttons {"Show me the instructions", "Do it for me"} default button 2 with title "WSTC"
			
			if the result is {button returned:"Do it for me"} then
				-- prompt the user to allow JS for AE
				tell application "System Events" to tell process "Safari"
					set isPresent to (exists menu bar item "Develop" of menu bar 1)
					-- if the dev menu isn't there, make it be there
					if not isPresent then toggle_Safari_dev_menu()
					-- click the proper menu item
					click menu item ¬
						"Allow JavaScript from Apple Events" of menu "Develop" of menu bar item "Develop" of menu bar 1
				end tell -- app "SE" --> process "Safari"
				
				-- this is super awkward but without it we just continue on, so...
				display dialog strWaitForJS buttons {"Continue"} default button 1 giving up after 600 with title "WSTC"
				
				-- if the dev menu wasn't there before, hide it
				if not isPresent then toggle_Safari_dev_menu()
				
			else -- i.e. user wants to see instructions
				-- open the instructions page
				tell application "Safari" to open location allowJSURL
				
				-- this is super awkward but without it we just continue on, so...
				display dialog strWaitForJS buttons {"Continue"} default button 1 giving up after 600 with title "WSTC"
			end if -- result from asking how user wants to allow JS access
			
		end try -- do JavaScript in Safari
	end if -- not initialized
	
	
end initialize




------------------------------------------------------
------------------------------------------------------
------ Import user settings from plist
------------------------------------------------------
------------------------------------------------------
to import_settings()
	-- import settings from the settings plist
	tell application "System Events"
		set s_plist to property list file (settingsPath)
		set email to value of property list item "email" of s_plist
		set encPin to value of property list item "pin" of s_plist
		set eventTitle to value of property list item "title" of s_plist
		set calName to value of property list item "calendar" of s_plist
		set address to value of property list item "address" of s_plist
		set populateDescription to value of property list item "descriptionIsShiftID" of s_plist
		set appendShift to value of property list item "appendShiftID" of s_plist
		set closeSafariWhenDone to value of property list item "closeSafariWhenDone" of s_plist
		set closeCalendarWhenDone to value of property list item "closeCalendarWhenDone" of s_plist
	end tell -- app "SE"
end import_settings




------------------------------------------------------
------------------------------------------------------
------ Solicit settings from user, save to plist
------------------------------------------------------
------------------------------------------------------
to create_settings()
	-- get info from user
	set eventTitle to my_prompt(titlePrompt, settingsPromptTitle, "Work")
	
	-- have user pick calendar from available calendars
	tell application "Calendar" to tell every calendar to set calNames to name
	set calName to ¬
		(choose from list calNames with prompt calPrompt with title settingsPromptTitle ¬
			default items item 1 of calNames OK button name "Continue")
	-- "choose from list" MUST have a "cancel" button, so if the user clicks it
	-- ...we default to the first calendar of the list
	if (calName = false) then
		set calName to item 1 of calNames
	else
		set calName to item 1 of calName
	end if
	
	-- get the rest of user settings
	set address to my_prompt(addressPrompt, settingsPromptTitle, "")
	set appendShift to my_prompt_bool(appendPrompt, settingsPromptTitle)
	set populateDescription to my_prompt_bool(populatePrompt, settingsPromptTitle)
	set closeSafariWhenDone to my_prompt_bool(closeSafariPrompt, settingsPromptTitle)
	set closeCalendarWhenDone to my_prompt_bool(closeCalendarPrompt, settingsPromptTitle)
	
	-- get login settings
	set email to my_prompt(emailPrompt, loginPromptTitle, "")
	set encPin to my_prompt_enc(pinPrompt, loginPromptTitle, pw) -- returns encrypted text
	
	-- create s_plist and save the info
	set progress completed steps to 16
	set progress additional description to "Saving user settings"
	tell application "System Events"
		-- set up the file
		set parentDict to make new property list item with properties {kind:record}
		set s_plist to make new property list file with properties {contents:parentDict, name:settingsPath}
		-- write back info
		tell property list items of s_plist
			make new property list item at end with properties {kind:string, name:"title", value:eventTitle}
			make new property list item at end with properties {kind:string, name:"calendar", value:calName}
			make new property list item at end with properties {kind:string, name:"address", value:address}
			make new property list item at end with properties {kind:boolean, name:"appendShiftID", value:appendShift}
			make new property list item at end with properties {kind:boolean, name:"descriptionIsShiftID", value:populateDescription}
			make new property list item at end with properties {kind:boolean, name:"closeSafariWhenDone", value:closeSafariWhenDone}
			make new property list item at end with properties {kind:boolean, name:"closeCalendarWhenDone", value:closeCalendarWhenDone}
			make new property list item at end with properties {kind:string, name:"email", value:email}
			make new property list item at end with properties {kind:string, name:"pin", value:encPin}
			make new property list item at end with properties {kind:record, name:"secQandA"}
		end tell -- prop list items of s_plist
	end tell -- app "SE"
	
end create_settings




------------------------------------------------------
------------------------------------------------------
------ Log in to WMT Scheduler
------------------------------------------------------
------------------------------------------------------
to login()
	-- open Web Scheduler splash page
	set progress additional description to "Opening WMT Scheduler login page"
	if not (safariWasRunning) then
		-- no need to activate, we activated it during initialize()
		-- open the WMT website in the current tab/window
		tell application "Safari" to set the URL of the front document to webScheduleURL
	else
		-- open the WMT website in a new tab/window (Safari will stay in background)
		tell application "Safari" to open location webScheduleURL
	end if -- not safariWasRunning
	
	-- wait until WMT splash page loads
	wait_for_page_load(strInWMTSplashPage)
	set progress completed steps to 22
	
	-- click the login button, which will take us to MyAccess
	set progress additional description to "Opening MyAccess login page"
	click_ID(wmtLoginButtonID)
	wait_for_page_load(strInEmailEntryPage)
	
	-- enter user's email address, repeating process if the email was rejected
	repeat
		-- enter email address and manually remove "disabled" attribute on button; submit
		set progress completed steps to 23
		set progress additional description to "Entering email address"
		input_by_ID(emailInputID, email)
		remove_attr_by_ID(emailButtonID, "disabled")
		set progress additional description to "Submitting email address"
		click_ID(emailButtonID)
		set progress completed steps to 24
		
		-- wait until PIN-and-question page loads
		if wait_for_page_load(strInPinEntryPage) then exit repeat
		
		-- if wfpl returns false we must still be on the email entry page
		-- so the email must have been wrong
		-- ask for user's corrected email and write it back to s_plist
		set email to my_prompt(badEmailPrompt, settingsPromptTitle, "")
		tell application "System Events" to tell property list items of s_plist to ¬
			make new property list item at end with properties {kind:string, name:"email", value:email}
	end repeat
	
	-- enter user's PIN and secret answer, repeating process if they were rejected
	repeat 4 times -- there are only three answers, no point in dragging it out
		
		set progress completed steps to 27
		set progress additional description to "Entering PIN"
		-- enter PIN (catching a decrypt error)
		try
			input_by_ID(pinInputID, str_dec(encPin, pw))
			delay 2 -- necessary to allow "signedChallenge" var to populate? I think the page "phones home" to get it
		on error
			-- an error means the saved decrypt PW is bad (script was recompiled)
			-- ask user to re-enter PIN
			set encPin to my_prompt_enc(badPassPrompt, loginPromptTitle, pw)
			
			-- enter the PIN
			input_by_ID(pinInputID, str_dec(encPin, pw))
			
			-- store the new encrypted pin		
			-- and because the password was bad, the secret questions are also no longer good
			tell application "System Events" to tell property list items of s_plist
				make new property list item at end with properties {kind:string, name:"pin", value:encPin}
				make new property list item at end with properties {kind:record, name:"secQandA"}
			end tell
			-- no need to delay—the user will have to enter their secret answer, giving "signedChallenge" time
		end try -- to enter PIN
		set progress completed steps to 28
		
		set progress additional description to "Entering secret question"
		-- determine secret question
		set question to get_val_by_selector("label", "for", answerInputID, 0, 0)
		
		-- attempt to look up answer using security question as key; if it isn't stored, prompt user
		try
			tell application "System Events" to tell property list item "secQandA" of s_plist ¬
				to tell property list item question to set encAnswer to value
		on error
			set encAnswer to my_prompt_enc(questionPrompt & question, loginPromptTitle, pw)
			tell application "System Events" to tell property list item "secQandA" of s_plist to ¬
				tell property list items to make new property list item at end ¬
					with properties {kind:string, name:question, value:encAnswer}
		end try -- get secret answer by question as key
		
		-- input the answer
		input_by_ID(answerInputID, str_dec(encAnswer, pw))
		set progress completed steps to 29
		
		set progress additional description to "Submitting PIN and secret question"
		-- I'm not sure how, but sometimes it happens that the PIN field is empty before the login button gets clicked
		-- this causes the page to issue a JavaScript alert and we can't interact with it!
		-- check now to make sure, and if it is re-enter it
		if (get_val_by_id(pinInputID) = "") then
			input_by_ID(pinInputID, str_dec(encPin, pw))
			delay 2 -- to allow "signedChallenge" to populate
		end if
		
		-- manually remove "disabled" attribute on button; submit
		remove_attr_by_ID(myAccessLoginButtonID, "disabled")
		click_ID(myAccessLoginButtonID)
		
		-- wait until home page loads
		if (wait_for_page_load(strInWorksheetView)) then exit repeat
		
		-- if wfpl returns false, MyAccess didn't like the PIN and/or secret answer
		-- ask user to re-enter PIN
		set encPin to my_prompt_enc(badPinPrompt, loginPromptTitle, pw)
		
		-- enter the PIN
		input_by_ID(pinInputID, str_dec(encPin, pw))
		
		-- store the new encrypted pin		
		-- it would be great to only remove the offending Q/A pair, but it's easier to just wipe them all
		tell application "System Events" to tell property list items of s_plist
			make new property list item at end with properties {kind:string, name:"pin", value:encPin}
			make new property list item at end with properties {kind:record, name:"secQandA"}
		end tell
		-- no need to delay—the user will have to enter their secret answer, giving "signedChallenge" time to populate
	end repeat
end login




------------------------------------------------------
------------------------------------------------------
------ Get current/upcoming pay period's shifts
------------------------------------------------------
------------------------------------------------------
to get_shifts()
	-- click link to go to "My Schedule" view, wait until page loads
	click_query_first("a", "href", scheduleViewURL)
	wait_for_page_load(strInScheduleView)
	set progress completed steps to 34
	
	set progress additional description to "Checking database file"
	-- check to make sure our database file exists; if not, create it
	tell application "System Events"
		if (exists file databasePath) then
			set db_plist to property list file databasePath
			set eventRec to property list item "eventIDs" of db_plist
			set shiftRec to property list item "shifts" of db_plist
		else -- i.e. if the file doesn't exist
			set parentDict to make new property list item with properties {kind:record}
			set db_plist to make new property list file with properties {contents:parentDict, name:databasePath}
			tell db_plist
				set eventRec to make new property list item at end with properties {kind:record, name:"eventIDs"}
				set shiftRec to make new property list item at end with properties {kind:record, name:"shifts"}
			end tell -- db_plist
		end if -- not exists file databasePath
	end tell -- app "SE"
	set progress completed steps to 35
	
	-- store the first Sunday of the pay period we defaulted to—this is the current pay period we're in
	set currentStartDate to date (get_val_by_selector("td", "align", "center", 0, indexOfDateInScheduleDiv))
	
	-- scrape one pay period's worth of shifts at a time and place them in the Settings plist
	repeat
		set progress additional description to "Scraping shifts"
		-- initialize list of shifts
		set shifts to {}
		
		-- get the 14 shifts one-by-one and add them to the new list
		repeat with thisDay from 0 to (daysInPP - 1)
			set progress additional description to "Scraping shifts (" & thisDay + 1 & " of " & daysInPP & ")"
			set thisShift to get_val_by_selector("td", "align", "center", thisDay, indexOfShiftInScheduleDiv)
			set shifts to (shifts & {thisShift})
		end repeat
		
		-- place shifts in the database plist as an array within "shifts", using the first date of the PP as the key
		set thisStartDate to get_val_by_selector("td", "align", "center", 0, indexOfDateInScheduleDiv)
		tell application "System Events" to tell property list items of shiftRec ¬
			to make new property list item at end with properties {kind:list, name:thisStartDate, value:shifts}
		
		-- check if the currently loaded pay period is the latest (last) available
		-- if not, select the next pay period
		tell application "Safari"
			tell document 1 to do JavaScript "
			var options = document.getElementsByTagName('option');
			var selOpt = document.querySelectorAll('option[selected]')[0];
			var i=0;
			while (i < options.length) { if (options[i] === selOpt) {break;} i++; }
			if ((options.length - i) != 1) {
				options[i].removeAttribute('selected');
				options[i+1].setAttribute('selected', 'selected');
				1 	// poor man's 'return' to AppleScript
			} else {
				0 	// poor man's 'return' to AppleScript
			}"
			if the result = 0 then exit repeat
			
			tell document 1 to do JavaScript "form1.submit();"
		end tell -- app "Safari"
		
		set progress completed steps to progress completed steps + 1
		-- if we're still here, it means we're waiting for the next PP to load
		wait_for_page_load(strInScheduleView)
	end repeat
	
end get_shifts




------------------------------------------------------
------------------------------------------------------
------ Enter the shifts as calendar events
------------------------------------------------------
------------------------------------------------------
to create_events()
	-- get a list of keys (PP start dates) for list of shifs
	tell application "System Events" to tell property list items of shiftRec to set startDates to name
	
	-- for each PP start date, make/check calendar events for all shifts
	repeat with startDate in startDates
		-- create lists of the assigned shifts and corresponding event IDs for this PP
		tell application "System Events"
			-- create a list of the assigned shifts for this PP
			tell property list items of property list item startDate of shiftRec to set shifts to value
			
			-- get a list of the corresponding event IDs for this PP, or initialize the list if they don't exist
			try
				tell property list items of property list item startDate of eventRec to set eventIDs to value
			on error
				set eventIDs to {}
				repeat daysInPP times
					set eventIDs to eventIDs & {""}
				end repeat
			end try
			
		end tell -- app "SE"
		
		-- get shift start and end times, compare to existing event (if any)
		repeat with dayNum from 1 to daysInPP
			set progress additional description to "Day " & dayNum & " of " & daysInPP
			
			-- get information for the current day's shift
			set thisDate to short date string of ((date startDate) + (dayNum - 1) * days)
			set thisShift to item dayNum of shifts
			set thisEventID to item dayNum of eventIDs
			set thisKey to thisShift -- we will strip thisShift down to a key, but want to save the full info as well
			set thisEventTitle to eventTitle
			
			-- initialize bool to indicate admin shift
			set isAdminShift to false
			
			-- make sure shift is a "true" work shift (not leave/RDO/training/etc)
			-- check for the abscence of "<" and the presence of a digit
			if ((not str_contains(thisKey, offsiteChar)) and (str_contains(thisKey, digits))) then
				-- check for admin shift indicator (@), log it, then get rid of it
				if str_contains(thisKey, adminChar) then
					set isAdminShift to true
					set thisKey to trim_first_chars(thisKey, 1)
				end if
				
				-- check for prepends (C, D, T, AREA), then get rid of them
				if str_contains(thisKey, oneCharPrepends) then set thisKey to trim_first_chars(thisKey, 1)
				if str_contains(thisKey, fourCharPrepends) then set thisKey to trim_first_chars(thisKey, 4)
				
				-- check for one-char appends ($, ^, !), then get rid of them
				if str_contains(thisKey, oneCharAppends) then set thisKey to trim_last_chars(thisKey, 1)
				
				-- check for other appends (OM, TMC), then get rid of them
				if str_contains(thisKey, twoCharAppends) then set thisKey to trim_last_chars(thisKey, 2)
				if str_contains(thisKey, threeCharAppends) then set thisKey to trim_last_chars(thisKey, 3)
				
				-- any character still left behind the number should be part of an actual shift key
				-- we use the key to get the shift start and end times
				-- an error means we couldn't get them (because the legends file is either missing or incomplete)
				-- so we run the get_shift_legends() handler and try again
				repeat 2 times
					try
						tell application "System Events" to tell property list item thisKey of property list file legendPath
							set startTime to value of property list item "start"
							set endTime to value of property list item "end"
						end tell
						exit repeat
					on error
						get_shift_legends()
					end try
				end repeat
				
				-- set up event start and end times
				set evtStartDate to date (thisDate & " " & startTime)
				set evtEndDate to date (thisDate & " " & endTime)
				
				-- check to see if shift spans midnight
				if (evtEndDate < evtStartDate) then set evtEndDate to (evtEndDate + 1 * days)
				
				-- if admin shift, add half an hour to the end time
				if (isAdminShift) then set evtEndDate to (evtEndDate + 30 * minutes)
				
				-- if user wanted, append shift name to title
				if (appendShift) then set thisEventTitle to (eventTitle & " " & thisShift)
				
				-- check to see if there is not already an existing event ID
				-- if there isn't, create the event; if there is, check event details and update details if necessary
				if (thisEventID = "") then
					-- create event
					tell application "Calendar" to tell calendar calName
						set thisEvent to make new event with properties ¬
							{summary:thisEventTitle, description:thisShift, start date:evtStartDate, end date:evtEndDate, location:address}
						if (populateDescription) then set description of thisEvent to thisShift
						
						-- store the event UID
						set item dayNum of eventIDs to (uid of thisEvent)
					end tell -- cal calName of app "Cal"
				else -- i.e. eventID is not empty
					-- edit existing event
					tell application "Calendar" to tell calendar calName
						-- make sure the event still exists
						if (exists event id thisEventID) then
							-- edit any part of the event that is not longer accurate
							set thisEvent to (event id thisEventID)
							tell thisEvent
								if (start date ≠ evtStartDate) then set start date to evtStartDate
								if (end date ≠ evtEndDate) then set end date to evtEndDate
								if (location ≠ address) then set location to address
								if (summary ≠ eventTitle) then set summary to thisEventTitle
								if (populateDescription) then if (description ≠ thisShift) then ¬
									set description to thisShift
							end tell -- thisEvent
						else -- i.e. if there is an ID in eventIDs but the event can't be found (user may have deleted it)
							-- go ahead and recreate it
							set thisEvent to make new event with properties ¬
								{summary:thisEventTitle, description:thisShift, start date:evtStartDate, end date:evtEndDate, location:address}
							if (populateDescription) then set description of thisEvent to thisShift
							
							-- event UID is read-only, so just get the new one and write it over the one we had stored
							set item dayNum of eventIDs to (uid of thisEvent)
						end if -- exists event id thisEventID
					end tell -- cal calName of app "Cal"
					
				end if -- thisEventID is not populated
				
			else -- i.e. shift is not a "true" work shift (RDO/leave/etc)
				-- check to see if there's an event ID for the day already—if so, delete event and blank out the entry
				set thisEventID to item dayNum of eventIDs
				if (thisEventID ≠ "") then
					-- check to make sure an event exists for us to delete, the user may have done it already
					tell application "Calendar" to tell calendar calName
						if (exists event id thisEventID) then ¬
							delete (event id thisEventID)
					end tell
					
					-- blank out the entry in eventIDs
					set item dayNum of eventIDs to ""
				end if -- thisEventID not empty
			end if -- shift is true work shift
			
		end repeat -- days in PP
		
		-- write out event IDs to s_plist for next time
		tell application "System Events" to tell property list items of eventRec to ¬
			make new property list item at end with properties ¬
				{kind:list, name:startDate, value:eventIDs}
	end repeat -- PPs saved in "shifts"
	
	-- refresh/reload calendars (i.e. upload new events or changes to cloud service)
	tell application "Calendar" to reload calendars
	
end create_events




------------------------------------------------------
------------------------------------------------------
------ Get shift legends (start/end times)
------------------------------------------------------
------------------------------------------------------
to get_shift_legends()
	set progress completed steps to 46
	set progress additional description to "Getting shift legends"
	
	-- click link to go to "Shift Legend" page, wait until page loads
	click_query_first("a", "href", shiftLegendURL)
	wait_for_page_load(strInLegendView)
	set progress completed steps to 47
	set progress additional description to "Getting shift legends (Preparing webpage)"
	
	-- strip spaces inside all centered divs on the page
	tell application "Safari" to tell document 1 to do JavaScript "
		var divs = document.querySelectorAll(\"div[align='center']\");
		for (var i=0; i < divs.length; i++) { divs[i].firstChild.replaceWith(divs[i].firstChild.nodeValue.replace(/\\s/g, '')); }"
	
	-- prepare plist contents
	tell application "System Events"
		set parentDict to make new property list item with properties {kind:record}
		set l_plist to make new property list file with properties {contents:parentDict, name:legendPath}
	end tell
	set progress completed steps to 48
	
	-- there are six centered header cells
	set thisDiv to numCenteredHeadingCells
	set i to 1
	
	-- we don't know how many shifts there are, so just keep going until we run out of divs and get an error
	-- luckily the only "div align=center" tags are the ones in this table—we won't get garbage at the end
	try
		repeat
			set progress additional description to "Getting shift legends (Shift " & i & ")"
			-- get the shift IDs and times
			set thisShift to get_val_by_selector("div", "align", "center", thisDiv, 0)
			set thisStartTime to get_val_by_selector("div", "align", "center", thisDiv + 1, 0)
			set thisEndTime to get_val_by_selector("div", "align", "center", thisDiv + 2, 0)
			
			-- there are six cells per row, but "Remarks" cell is not centered
			set thisDiv to (thisDiv + numCenteredContentCells)
			
			-- write shift IDs and times into s_plist
			tell application "System Events"
				tell property list items of l_plist to ¬
					make new property list item at end ¬
						with properties {kind:record, name:thisShift}
				tell property list item thisShift of l_plist
					make new property list item at end ¬
						with properties {kind:string, name:"start", value:thisStartTime}
					make new property list item at end ¬
						with properties {kind:string, name:"end", value:thisEndTime}
				end tell -- prop list item thisShift
			end tell -- app "SE"
			
			set i to i + 1
		end repeat
	on error
		-- do nothing; an error means we reached the end of the list of shifts
	end try
	
	set progress additional description to ""
	
end get_shift_legends



------------------------------------------------------
------------------------------------------------------
------ Close out the applications we used
------------------------------------------------------
------------------------------------------------------
to close_apps()
	if (closeSafariWhenDone) then
		if (safariWasRunning) then
			tell application "Safari" to close current tab of window 1
		else
			tell application "Safari" to quit
		end if
	end if
	
	if (closeCalendarWhenDone) then
		if (calendarWasRunning) then
			tell application "Calendar" to close window 1
		else
			-- we want to let Cal stay open for a moment to at least attempt a cloud sync
			-- and strangely "delay 5" doesn't do it—the app quits immediately
			-- this doesn't seem to work either though
			tell application "System Events" to do shell script "/bin/sleep 5"
			tell application "Calendar"
				do shell script "sleep 5"
				quit
			end tell
		end if
	end if
end close_apps




------------------------------------------------------
------------------------------------------------------
------ Remove old info from the database
------------------------------------------------------
------------------------------------------------------
to clean_up_db()
	-- remove old pay periods to inhibit file bloat
	-- unfortunately the "delete" command has no effect on plist items
	-- ...so we have to re-create the records and populate only with good PPs
	set savedPPs to {}
	tell application "System Events" to tell property list items of shiftRec to set startDates to name
	repeat with startDate in startDates
		-- we want to keep only startDates that are less than one PP old compared to the current PP
		-- i.e. only keep the current PP and newer
		if (currentStartDate - (date startDate)) < (1 * daysInPP * days) then ¬
			tell application "System Events" to set savedPPs to savedPPs & {{startDate, ¬
				(value of property list item startDate of shiftRec), ¬
				(value of property list item startDate of eventRec)}}
	end repeat -- startDate in startDates
	
	tell application "System Events"
		-- wipe out the records by recreating them
		tell property list items of db_plist
			make new property list item at end with properties {kind:record, name:"eventIDs"}
			make new property list item at end with properties {kind:record, name:"shifts"}
		end tell
		
		-- populate the new records with the saved PPs
		repeat with i from 1 to (length of savedPPs)
			tell property list items of shiftRec to make new property list item at end ¬
				with properties {kind:list, name:(item 1 of item i of savedPPs), value:(item 2 of item i of savedPPs)}
			tell property list items of eventRec to make new property list item at end ¬
				with properties {kind:list, name:(item 1 of item i of savedPPs), value:(item 3 of item i of savedPPs)}
		end repeat -- savedPP in savedPPs
	end tell -- app "SE"
	
end clean_up_db




------------------------------------------------------
------------------------------------------------------
------ Non-main handlers
------------------------------------------------------
------------------------------------------------------

-- simulate button press on HTML element, found by tag's ID
to click_ID(tagID)
	tell application "Safari" to tell document 1 to do JavaScript ¬
		"document.getElementById('" & tagID & "').click();"
end click_ID

-- simulate button press on HTML element, found by first tag with specified attribute
to click_query_first(tag, attr, attrVal)
	tell application "Safari" to tell document 1 to do JavaScript ¬
		"document.querySelectorAll('" & tag & "[" & attr & "=\"" & attrVal & "\"]')[0].click()"
end click_query_first

-- get value of an HTML tag, found by tag's ID
to get_val_by_id(tagID)
	tell application "Safari" to tell document 1 to do JavaScript ¬
		"document.getElementById('" & tagID & "').value"
end get_val_by_id

-- get value of an HTML tag, found by specified tag with specified attribute
to get_val_by_selector(tag, attr, attrVal, tagIndex, nodeIndex)
	tell application "Safari" to tell document 1 to return (do JavaScript ¬
		"document.querySelectorAll('" & tag & "[" & attr & "=\"" & attrVal & "\"]')[" & tagIndex & ¬
		"].childNodes[" & nodeIndex & "].nodeValue")
end get_val_by_selector

-- enter text into an HTML input field, found by tag's ID
to input_by_ID(tagID, theText)
	tell application "Safari" to tell document 1 to do JavaScript ¬
		"document.getElementById('" & tagID & "').value = '" & theText & "';"
end input_by_ID

-- activate an application, then return the original process to the front
to my_activate_return(appName, makeHidden)
	tell application "System Events" to set curProc to first process whose frontmost is true
	tell application appName
		activate
		if (makeHidden) then set visible of every window to false
	end tell
	tell application "System Events" to set frontmost of curProc to true
end my_activate_return

-- wrapper for "display dialog" to get yes/no answer
to my_prompt_bool(promptText, titleText)
	set res to false
	set progAddtlOld to progress additional description
	set progress additional description to strWaitForUser
	display dialog promptText with title titleText buttons {"No", "Yes"} default button 2
	if the result is {button returned:"Yes"} then set res to true
	set progress additional description to progAddtlOld
	return res
end my_prompt_bool

-- wrapper for "display dialog" to get text input
to my_prompt(promptText, titleText, defAns)
	set progAddtlOld to progress additional description
	set progress additional description to strWaitForUser
	display dialog promptText with title titleText default answer defAns buttons {"Continue"} default button 1
	set r to text returned of the result
	set progress additional description to progAddtlOld
	return r
end my_prompt

-- wrapper for "display dialog" to get text input with hidden answer, returning the encrypted result
to my_prompt_enc(promptText, titleText, pw)
	set progAddtlOld to progress additional description
	set progress additional description to strWaitForUser
	display dialog promptText with title titleText default answer "" buttons {"Continue"} default button 1 with icon caution with hidden answer
	set r to str_enc(text returned of the result, pw)
	set progress additional description to progAddtlOld
	return r
end my_prompt_enc

-- remove attribute from HTML tag, found using tag's ID
to remove_attr_by_ID(tagID, attr)
	tell application "Safari" to tell document 1 to do JavaScript ¬
		"document.getElementById('" & tagID & "').removeAttribute('" & attr & "');"
end remove_attr_by_ID

-- check if string contains any of the submitted substrings
-- Darrick Herwehe, https://stackoverflow.com/a/43783426, adapted
to str_contains(str, matchStrs)
	repeat with thisSubStr in matchStrs
		if str contains thisSubStr then return true
	end repeat
	return false
end str_contains

-- decrypt a base64 string with a password
to str_dec(str, pass)
	return do shell script "echo " & str & " | openssl aes-256-cbc -a -A -d -k " & pass
end str_dec

-- encrypt a string with a password (result in base64)
to str_enc(str, pass)
	return do shell script "echo " & quoted form of str & " | openssl aes-256-cbc -a -A -k " & pass
end str_enc

-- toggle the "Develop" menu bar menu in Safari
to toggle_Safari_dev_menu()
	tell application "System Events" to tell process "Safari"
		click menu item "Preferences…" of menu "Safari" of menu bar item "Safari" of menu bar 1
		click button "Advanced" of toolbar 1 of window 1
		click checkbox "Show Develop menu in menu bar" of group 1 of group 1 of window 1
		click button 1 of window 1
	end tell
end toggle_Safari_dev_menu

-- trim arbitrary number of chars from front of string
-- if chars-to-trim is more than chars-in-string, returns empty string
to trim_first_chars(str, n)
	set chars to characters of str
	set newchars to chars
	repeat with i from 1 to n
		repeat with j from 2 to (count chars)
			set item (j - 1) of newchars to item j of chars
		end repeat
		set last item of newchars to ""
	end repeat
	return newchars as string
end trim_first_chars

-- trim arbitrary number of chars from end of string
to trim_last_chars(str, n)
	set str to trim_first_chars((reverse of (characters of str)) as string, n)
	return (reverse of (characters of str) as string)
end trim_last_chars

-- wait for a webpage to load
-- I'm not very happy with this handler
-- but there doesn't seem to be a truly robust working solution, let alone an elegant one
-- because there are two distinct possibilities:
--   1) the new page loads and contains the desired text from the correct page (yay!)
--   2) the new page is actually the same as the old page, or the wrong page loads (boo!)
-- so the "compromise" method of waiting until the page source contains the desired text
-- is liable to hang for several seconds while it waits for a condition that will never be true
--
-- methods based on waiting until the page starts loading, then waiting until it finishes
-- (like waiting until source = "", or until readyState = "loading", or until source ≠ oldSource)
-- may fail: the page may pass those states before the test is evaluated
-- or the state may exist for such a short time that we miss it
--
-- the optimal solution, as described at https://apple.stackexchange.com/a/343633,
-- is to use System Events to check the state of the "reload/stop load" button
-- (button changes state as soon as the request is sent, and remains until 100% rendered)
-- but even reading this button requires the user to allow SE to control the system
-- and I don't want to ask for that
to wait_for_page_load(desiredText)
	delay 0.5
	delay globalPreDelay
	repeat 2 times
		tell application "Safari"
			repeat 30 times -- six seconds
				if (document 1's source contains desiredText) then exit repeat
				delay 0.2
			end repeat
			
			tell document 1 to do JavaScript "document.readyState"
			repeat until (the result = "complete")
				delay 0.2
				tell document 1 to do JavaScript "document.readyState"
			end repeat
			
			if document 1's source contains desiredText then return true
		end tell
	end repeat
	return false
end wait_for_page_load
