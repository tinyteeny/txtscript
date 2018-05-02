On Error Resume Next
mname = InputBox("Enter Machine Name", "Reboot Machine")
If Len(mname) = 0 Then Wscript.Quit

if Msgbox("Are you sure you want to reboot machine " & mname, vbYesNo, "Reboot Machine") = vbYes then

        Set OpSysSet = GetObject("winmgmts:{impersonationLevel=impersonate,(RemoteShutdown)}//" & mname).ExecQuery("select * from Win32_OperatingSystem where Primary=true")
        for each OpSys in OpSysSet
            OpSys.Reboot()
         next
end if
