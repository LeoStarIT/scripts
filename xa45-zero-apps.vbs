set farm = CreateObject("MetaFrameCOM.MetaFrameFarm")
farm.Initialize(1)
for each server in farm.Servers
if server.Applications.Count = 0 then
WScript.Echo server.ServerName & " " & server.Applications.Count
end if
next