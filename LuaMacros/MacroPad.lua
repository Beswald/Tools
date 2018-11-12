local InitMode = "TRUE";
local DynamicAssign = "TRUE";

LogScriptMsg = function(key,logMsg)
	local file = io.open("LuaKeypadLog.txt", "a") -- writing this string to a text file on disk is probably NOT the best method. Feel free to program something better!
		
        file:write(os.date("%c"))
        
		if (key == -1) then
			file:write(" - ",logMsg,"\n");
		else			
			file:write(" - ",logMsg,"(",key,")","\n")			
		end
		
		file:flush() --"flush" means "save"
		file:close()
end

HandleMacroKey = function(key)
	if(key == 111) then         -- H Drive
		lmc_spawn("wscript.exe","Scripts\\OpenDir.vbs","TABLE","0")
	elseif(key == 106) then     -- GitLocal
		lmc_spawn("wscript.exe","Scripts\\OpenDir.vbs","TABLE","1")
	elseif(key == 8) then       -- ProdDev
		lmc_spawn("wscript.exe","Scripts\\OpenDir.vbs","TABLE","2")
	elseif(key == 97) then		-- Rename (Generic)
		lmc_send_keys('{F2}')
	elseif(key == 35) then		-- Admin Command Prompt
		lmc_spawn("wscript.exe","Scripts\\OpenAdminCmd.vbs")
	elseif(key == 98) then		-- Open File in Notepad++
		lmc_spawn("wscript.exe","Scripts\\OpenFileInNotepad.vbs")
	elseif(key == 40) then		-- Open Notepad++ Blank
		lmc_spawn("wscript.exe","Scripts\\OpenFileInNotepad.vbs", "BLANK")
	elseif(key == 99) then		-- Toggle Mute
		lmc_send_input(173, 0, 0)  -- 0xAD 
	elseif(key == 34) then		-- Show Last TimeTrack Entry
		lmc_spawn("wscript.exe","Scripts\\LastTimeTrackEntry.vbs")
	elseif (key == 100) then	-- Navigate Backward (VS)
		lmc_send_keys('^-')                  
	elseif (key == 101) then	-- Show Design (VS)
		lmc_send_keys('+{F7}')
	elseif (key == 102) then	-- Show Code (VS)
		lmc_send_keys('{F7}')
	elseif (key == 103) then	-- Find All (VS)
		lmc_send_keys('^kr')
	elseif (key == 104) then	-- Goto Definition (VS)
		lmc_send_keys('{F12}')
	elseif (key == 105) then	-- Toggle Outlining (VS)
		lmc_send_keys('^ml')
	elseif (key == 33) then		-- Close All Windows (VS)
		lmc_spawn("wscript.exe","Scripts\\CloseAllWindows.vbs")		
	elseif (key == 107) then	-- Launch VS
		lmc_spawn("C:\\Program Files (x86)\\Microsoft Visual Studio\\2017\\Professional\\Common7\\IDE\\devenv.exe")
	elseif (key == 96) then		-- Lock PC	
		lmc_spawn("wscript.exe","Scripts\\LockWorkstation.vbs")
		LogScriptMsg(-1,"Workstation Locked")
	elseif (key == 45) then		-- Restart PC	
		lmc_spawn("wscript.exe","Scripts\\RestartPC.vbs")
		lmc_spawn("wscript.exe","Scripts\\TimeTrack.vbs", "OUT")
        LogScriptMsg(-1,"Workstation Restarted")
	elseif (key == 46) then		-- Task Manager
		lmc_spawn("wscript.exe","Scripts\\TaskMan.vbs")        
	elseif (key == 109) then	-- Time Tracking
		lmc_spawn("wscript.exe","Scripts\\TimeTrack.vbs")
	elseif (key == 110) then	-- Switch Windows
		--lmc_send_keys('^(#{TAB})')
		lmc_send_keys('^(%{TAB})')
	elseif (key == 13) then		-- Enter
		lmc_send_keys('{ENTER}')
	elseif (key == 144) then	-- NumLock
		LogScriptMsg(-1,"NumLock Toggled")	
	else
		LogScriptMsg(key,"Undefined Key Press")
	end
end

local MacroPad = ""

if(DynamicAssign == "TRUE") then
	lmc_assign_keyboard('MACRO_PAD')
	
	LogScriptMsg(-1,"Macro Startup, Macro Keypad Assigned Dynamically")
else
	MacroPad = lmc_device_set_name('MACRO_PAD', '369FC0E4')  
	-- Get the above value from device manager. Looks like this : HID\VID_1A2C&PID_4224&MI_00\9&369FC0E4&0&0000
	
	LogScriptMsg(-1,"Macro Startup, Keypad Assigned - 369FC0E4")	
	lmc_spawn("wscript.exe","Scripts\\TimeTrack.vbs", "IN")
	InitMode = "FALSE"
end

lmc.minimizeToTray = true
lmc_minimize()

-- define callback for whole device
lmc_set_handler('MACRO_PAD', function(button, direction)

	WinTitle=lmc_get_window_title()

	if (string.match(WinTitle, "Visual Studio")) then
		if (direction == 1) then return end  -- ignore down.   (Volume Down/Up are View Code/View Design when Visual Studio is active window)
	else 
		if(button == 101) then  -- Volume Control
			if (direction == 1) then -- Act on down to get repeat, up is ignored
                lmc_send_input(174, 0, 0)  -- 0xAE	-- Volume Down
			end
		elseif(button == 102) then
			if (direction == 1) then -- Act on down to get repeat, up is ignored
				lmc_send_input(175, 0, 0)  -- 0xAF	-- Volume Up
			end
		else
			if (direction == 1) then return end  -- ignore down, act on up (i.e. Full Cycle of Key)   
		end
	end

	if(InitMode == "TRUE") then   -- Throw away first keypress when assigning input device
		InitMode = "FALSE"
		LogScriptMsg(-1,"Init Mode Handled (MacroPad)")
		lmc_spawn("wscript.exe","Scripts\\TimeTrack.vbs", "IN")		
	else
		HandleMacroKey(button)
	end
end)


