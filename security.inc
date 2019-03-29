<%
sub CopyNTFSSecuritySettings(byval SourceObjectName, byval TargetObjectName)
  on error resume next

  dim SourceWMIPath, TargetWMIPath
  dim WMI
  dim SecurityDescriptor
  dim SourceSecuritySettings, TargetSecuritySettings

  SourceWMIPath = Replace(SourceObjectName, "\", "\\")
  TargetWMIPath = Replace(TargetObjectName, "\", "\\")

  set WMI = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & "." & "\root\CIMv2")
  set SourceSecuritySettings = WMI.Get("Win32_LogicalFileSecuritySetting.Path='" & SourceWMIPath & "'")

  if SourceSecuritySettings.GetSecurityDescriptor(SecurityDescriptor) = 0 then
     set TargetSecuritySettings = WMI.Get("Win32_LogicalFileSecuritySetting.Path='" & TargetWMIPath & "'")
     TargetSecuritySettings.SetSecurityDescriptor(SecurityDescriptor)
     set TargetSecuritySettings = Nothing
  end if

  set SecurityDescriptor = Nothing
  set SourceSecuritySettings = Nothing
  set WMI = Nothing
end sub
%>
