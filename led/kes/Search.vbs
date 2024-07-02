dim ws
Set ws=CreateObject("wscript.shell")
for each x in getobject("winmgmts:").execquery("select * from CIM_DataFile where FileName = 'ThunderStart' and Extension = 'exe'")
str = x.name
ws.run chr(34) & str & chr(34)
next