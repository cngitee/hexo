On Error Resume Next
Dim strComputer,GOj,Wsh,fso,oReg,Datad_a,Datad_b,Datad_c,Location,Location1,strKeyPath_1,strKeyPath_2,strKey,datd,Itemss,Rt,objFile,ID1,Arrtr,Items_datc,fpcth,fcy,i,Fname,F,protected,datd1,o,z,p,g
strComputer = "."
Set GOj = GetObject("winmgmts:")
Set Wsh = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
Const HKLM = &H80000002
strKeyPath_1 = "SOFTWARE":strKeyPath_2 = "SOFTWARE\Wow6432Node"
strKey = "SOFTWARE\Microsoft\SystemCertificates\SPC\Certificates"
if Err.Number <> 0 then
msgbox "权限不足。请以管理员身份运行, 或登陆“Administrator”帐户再执行此操作！"  ,48,"警告"
WScript.Quit
end if
Function LiRu(strKeyPath,pro)
On Error Resume Next
arrSubKeys = "":subkey = ""
oReg.EnumKey HKLM, strKeyPath& "\KasperskyLab"& pro, arrSubKeys
For Each subkey In arrSubKeys 
Err.Clear
Datad_c = Wsh.RegRead("HKLM"& "\"& strKeyPath& "\KasperskyLab"& pro& "\"& subkey &"\environment\ProductRoot")
if Err.Number = 0 and fso.FileExists(Datad_c& "\avp.exe") then
Err.Clear
Wsh.RegRead("HKLM\"&  strKeyPath& "\KasperskyLab"& pro& "\LicStrg\")
if Err.Number = 0 then Location1 = strKeyPath& "\KasperskyLab"& pro& "\LicStrg"
Err.Clear
Wsh.RegRead("HKLM\"&  strKeyPath& "\KasperskyLab"& pro& "\LicStorage\")
if Err.Number = 0 then Location1 = strKeyPath& "\KasperskyLab"& pro& "\LicStorage"
Location = strKeyPath& "\KasperskyLab"& pro& "\"& subkey
end if
Next
End Function
Err.Clear
Location = ""
protected = "\protected"
Call LiRu(strKeyPath_2,protected):Call LiRu(strKeyPath_1,protected)
protected = ""
Call LiRu(strKeyPath_2,protected):Call LiRu(strKeyPath_1,protected)
Datad_a = "":Datad_a = Wsh.RegRead("HKLM"& "\"& Location &"\environment\DataRoot")
Datad_b = "":Datad_b = Wsh.RegRead("HKLM"& "\"& Location &"\environment\ProductType")
Datad_c = "":Datad_c = Wsh.RegRead("HKLM"& "\"& Location &"\environment\ProductRoot")
bt1 = ""
for k=1 to Len(Datad_b)
bt1 = bt1& hex(AscW(Mid(Datad_b,k,1)))& "00"
next
Err.Clear
fpcth=wscript.arguments(0)
oReg.EnumKey HKLM, strKey, arrValues
For i=0 To UBound(arrValues)
p = 0:p = Wsh.RegRead("HKLM"& "\"& strKey& "\"& arrValues(i)& "\BlobCount")
if p > 1 then
datd = ""
for n=0 to p-1
oReg.GetBinaryValue HKLM, strKey &"\"& arrValues(i)& "\Blob"& n,"Blob",strValues
for b=0  to UBound(strValues)
datd = datd& Right("0"& Hex(strValues(b)), 2)
next
next
else
oReg.GetBinaryValue HKLM, strKey &"\"& arrValues(i),"Blob",strValue
datd = ""
for b = 0 to UBound(strValue)
datd = datd& Right("0"& Hex(strValue(b)), 2)
next
end if
if Instr(datd, bt1) > 0 and bt1 <> "" then
if Instr(Mid(datd,Instr(datd, bt1)+Len(bt1)), bt1) > 0 then
datd1 = Mid(datd, Instr(1,datd,"2000000001000000",1))
Itemss = arrValues(i)
end if
end if
Err.Clear
Wsh.RegRead("HKLM"& "\"& Location1& "\")
if  Err.Number = 0  and fpcth = "" then
Select Case msgbox( "确定备份「卡巴斯基」的授权吗 ？ ",68, "提示")
Case 7 WScript.Quit 
'执行删除自身
'Set fso = CreateObject("Scripting.FileSystemObject")
'f = fso.DeleteFile(WScript.ScriptName) '执行完成
end Select
Fname = fso.GetParentFoldername(Wscript.scriptfullname)& "\kaspersky_"& Datad_b& ".dat"
'P = createobject("Scripting.FileSystemObject").GetFolder(".").Path '当前目录
' msgbox "首次备份在" & P & "\kaspersky_kes.dat",48,"提示"

'Const OverwriteExisting=True
'Set objFSO=CreateObject("Scripting.FileSystemObject")
D=year(Now)&"_"&Month(Now)&"_"&day(Now) '取日期
'P = createobject("Scripting.FileSystemObject").GetFolder(".").Path '当前目录
'objFSO.CopyFile P & "\kaspersky_kes.dat","D:\kaspersky_kes" & D &".dat", OverwriteExisting '备份D盘
msgbox "D:\kaspersky_kes"& D &".dat",68,"成功备份"

F = 2
if fso.FileExists(Fname) then
G = msgbox ("备份文件“kaspersky_"& Datad_b& ".dat"& "”已存在，是否覆盖 ？ ", 68, "警告")
if G = 2 then F = 1
end if
oReg.GetBinaryValue HKLM, Location1, Datad_b, strValue1
datdn = ""
for b = 0 to UBound(strValue1)
datdn = datdn& Right("0"& Hex(strValue1(b)), 2)
next
Set RS=createObject("ADODB.Recordset")
L=Len(datdn)/2:RS.Fields.Append "m",205,L:RS.Open:RS.AddNew:RS("m")=datdn&ChrB(0):RS.update:datdn=RS("m").GetChunk(L)
with createObject("ADODB.Stream"):.Mode = 3:.Type = 1:.Open():.Write datdn:.SaveToFile Fname,F:end with
WScript.Quit
end if
if Right(Left(datd,9),7) = "A700000"  and fpcth <> "" then
Select Case msgbox( "确定为「卡巴斯基」添加此授权吗 ？ ",68, "提示")
Case 7 WScript.Quit
end Select
Err.Clear
if p > 1 then
for n=0 to p-1
Wsh.RegDelete "HKLM\"& strKey &"\"& arrValues(i)& "\Blob"& n& "\"
next
end if
Wsh.RegDelete "HKLM\"& strKey &"\"& arrValues(i)&"\"
if Err.Number <> 0 then   
msgbox "权限不足。请以管理员身份运行, 或登陆“Administrator”帐户再执行此操作！"  ,48,"警告"
WScript.Quit
end if
Err.Clear
Wsh.RegWrite "HKLM"& "\"& Location &"\settings\Ins_InitMode", "1", "REG_DWORD"
if Err.Number <> 0 then   
msgbox "请关闭卡巴斯基的自我保护再执行此操作！"  ,48,"警告"
WScript.Quit
end if
Wsh.RegWrite "HKLM"& "\"& Location &"\settings\EnableSelfProtection", "1", "REG_DWORD"
Wsh.RegDelete  "HKLM"& "\"& Location &"\watchdog\LicenseInfo\"
fso.DeleteFile(Datad_a& "\Data\stor_"& Datad_b& ".bin" )
with createobject("adodb.stream") 
.type=1:.open:.loadfromfile fpcth:str=.read:sl=lenb(str) 
end with
if ".lic" = Right(fpcth,4) then
for b=1 to sl 
bt=ascb(midb(str,b,1)) 
objFile = objFile& Right("0"& Hex(bt-18), 2)
next 
elseif ".dat" = Right(fpcth,4) then
for b=1 to sl 
bt=ascb(midb(str,b,1)) 
objFile = objFile& Right("0"& Hex(bt), 2)
next
else
msgbox "不能识别无法加载的文件！"  ,48,"警告 "
WScript.Quit
end if
datdg=Mid(objFile, Instr(1,objFile,"4B4C737700004B4C",1))
Redim Items_d(len(datdg)/2-1):fcy = 0
for h=1  to len(datdg) step 2 
Items_d(fcy) = "&H"& Mid(Trim(datdg), h, 2)
fcy = fcy+1
next
n = len("000000"& hex(fcy))
for b=1 to 4
o = o & Mid("000000"& hex(fcy), n-1, 2) 
n=n-2
next 
oReg.SetBinaryValue HKLM,Location1, Datad_b, Items_d
objFile1 = "10A7000001000000"& o& Trim(datdg)& Mid(datd, Instr(1,datd,"0300000001000000",1),64)& datd1
Redim Items_datc(len(objFile1)/2-1):fcy = 0
for h=1  to len(objFile1) step 2 
Items_datc(fcy) = "&H"& Mid(objFile1, h, 2)
fcy = fcy+1
next
if Itemss <> "" and Datad_b <> "" and datd1 <> "" then
oReg.CreateKey HKLM, strKey &"\"& Itemss
if fcy < 12289 then
oReg.SetBinaryValue HKLM,strKey &"\"& Itemss, "Blob", Items_datc
else
Wsh.RegWrite "HKLM"& "\"& strKey& "\"& Itemss& "\BlobCount", fcy\12288+1, "REG_DWORD"
Wsh.RegWrite "HKLM"& "\"& strKey& "\"&  Itemss& "\BlobLength",fcy, "REG_DWORD"
Redim Items_datd(12287)
r =0
for g = 0 to fcy
Items_datd(r) = Items_datc(g)
r =r+1
if r = 12288 then  
r = 0
oReg.CreateKey HKLM, strKey &"\"& Itemss&"\"& "Blob"&  g\12288
oReg.SetBinaryValue HKLM,strKey &"\"& Itemss&"\"& "Blob"&  g\12288, "Blob", Items_datd
''Redim Items_datd(12287)
if fcy - g < 12288  then  Redim Items_datd(fcy - g-2)
end if
next
oReg.CreateKey HKLM, strKey &"\"& Itemss&"\"& "Blob"&  g\12288
oReg.SetBinaryValue HKLM,strKey &"\"& Itemss&"\"& "Blob"&  g\12288, "Blob", Items_datd
end if
end if
end if
Next
if fpcth <> "" then
if Itemss = "" or Datad_b = "" or datd1 = "" then
oReg.EnumKey HKLM, strKey, arrValues1
For V=0 To UBound(arrValues1)
Wsh.RegDelete "HKLM\"& strKey &"\"& arrValues1(V)&"\"
NEXT
msgbox " 未找到数据。 请重新启动一遍“卡巴斯基”后再尝试加载授权。  "  ,64,"警告"
WScript.Quit 
end if
Arrtr = "" 
Set ID1 = GOj.ExecQuery("select * from win32_process where name like 'avp.exe'" )
for Each i In ID1
Arrtr = Arrtr & i.name
next
if Len(Arrtr) > 0 then
msgbox "授权加载完成； 手动重启“卡巴斯基”生效 。"  ,64,"警告"
else
Wsh.Run  Chr(34) & Datad_c &"\avp.com"& Chr(34),true,true
Wsh.Run  Chr(34) & Datad_c &"\avp.com"& Chr(34),true
Wsh.Run  Chr(34) & Datad_c &"\avpui.exe"& Chr(34),true
end if
end if
Set GOj = Nothing:Set Wsh = Nothing:Set fso = Nothing:Set oReg = Nothing
Dim file,fileName '删除授权1
Set file=CreateObject("Scripting.FileSystemObject")
filePath = createobject("Scripting.FileSystemObject").GetFolder(".").Path
file.DeleteFile(filePath &"\" & "kaspersky_kes1.dat")
Set fso = CreateObject("Scripting.FileSystemObject") '删除自身
f = fso.DeleteFile(WScript.ScriptName)
dim WSHshell '名称不能重复 结束进程
set WSHshell = wscript.createobject("wscript.shell") 
WSHshell.run "taskkill /im wscript.exe /f ",0 ,true 
WScript.Quit '退出
