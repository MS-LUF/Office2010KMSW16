' updated by Orange to fix MS Issue by Lucas Cueff
' Copyright (c) Microsoft Corporation. All rights reserved.
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
CONST wshOK                =0
CONST VALUE_ICON_WARNING        =16
CONST wshYesNoDialog            =4
CONST VALUE_ICON_QUESTIONMARK        =32
CONST VALUE_ICON_INFORMATION        =64
CONST WindowsAppId                      = "55c92734-d682-4d71-983e-d6ec3f16059f"
CONST OfficeKmsSkuid            = "bfe7a195-4f8f-4f0b-a622-cf13c7d16864"
CONST VALUE_APPNAME             = "Microsoft Office 2010 KMS Host License Pack"
Const LocResource            = "kms_host.xml"
CONST MSG_SLUIEXE            = "slui.exe 0x2a 0x"
CONST MSG_RESOURCE_NONE            = "File not found: "
CONST FlavorEnterprise            = "ENTERPRISE"
CONST FlavorStandard            = "STANDARD"
CONST FlavorDataCenter            = "DATACENTER"
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Set WshShell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
workingDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
getEngine()

Set objWMIService = GetObject("winmgmts:" _
                & "{impersonationLevel=impersonate}!\\" & "." & "\root\cimv2")
Dim globalValue, tmpValue, tmpValue1, isCmdline
isCmdline = False

Select Case WSCript.Arguments.Count
    Case 0
    Case 1
        var1 = WSCript.Arguments(0)
        isCmdline = True
End Select

Call Main(var1)
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Sub Main(strProductKey)
    
On Error Resume Next

verifyFileExists LocResource

folder = "unknown"
For Each objOS in GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")
    Ver = Split(objOS.Version, ".", -1, 1)
    For counter=0 To uBound(Ver)
		Ver(counter) = CInt(Ver(counter))
	Next
    ' Win2K3
    If (Ver(0) = "5" And Ver(1) = "2" And (objOS.ProductType = 2 Or objOS.ProductType = 3)) Then
        'Check for supported OS flavor
        intPosition = InStr(UCase(objOS.Caption),FlavorEnterprise)
        If intPosition = 0 Then intPosition = InStr(UCase(objOS.Caption),FlavorStandard)
        If intPosition = 0 Then intPosition = InStr(UCase(objOS.Caption),FlavorDataCenter)
        If intPosition <> 0 Then folder = "win2k3" : CheckSPP(objWMIService) : Exit For
    End If
        
    ' win7/r2
    If (Ver(0) = "6" And Ver(1) = "1") Then
       folder = "win7"
       Exit For
    End If
    
    'patch
	' win8 or greater
    If (Ver(0) = "6" And Ver(1) >= "2") Or (Ver(0) >= "7") Or (Ver(0) = "10") Then
	' end of patch
       folder = "win8"
       Exit For
    End If
Next

If folder = "unknown" Then
    getResource "MSG_UNSUPPOS"
    WshShell.Popup globalValue,,VALUE_APPNAME, wshOK + VALUE_ICON_WARNING
    pauseExit()
End If

For Each objService in objWMIService.InstancesOf("SoftwareLicensingService")
    Set objSpp = objService
    Exit For
Next

Set productinstances = objWMIService.InstancesOf("SoftwareLicensingProduct")

Select Case folder
    Case "win7", "win8"
        isVL = False
        For Each instance in productinstances
            If (LCase(instance.ApplicationId) = WindowsAppId) Then
                intOccur = InStr(UCase(instance.Description),"VOLUME_KMS")
                If intOccur <> 0 Then
                    intOccurClient = InStr(UCase(instance.Description),"VOLUME_KMSCLIENT")
                    If intOccurClient = 0 Then
                        isVL = True
                        Exit For
                    End If
                End If            
            End If
        Next
        
        If isVL = False Then
            getResource "MSG_WINRET"
            WshShell.Popup globalValue,,VALUE_APPNAME, wshOK + VALUE_ICON_WARNING
            pauseExit()
        End If
    Case Else
End Select
 
errHandle("")

Set objFolder = fso.GetFolder(workingDir & folder)

If Not fso.FolderExists(objFolder) Then
    getResource "MSG_NODIR"
    WshShell.Popup globalValue,,VALUE_APPNAME, wshOK + VALUE_ICON_WARNING
    pauseExit()
End If

getResource "MSG_INSTALLLICENSES"
WScript.Echo globalValue
WScript.Echo vbcr
Err.Clear()

Set fileList = objFolder.Files
For Each objFile in fileList
    fileName = objFile.name
    If 0 <> InStr(fileName, "xrm-ms") Then
        InstallKMSLicense objSpp, workingDir & folder & "\" & fileName, folder
    End If
Next

WScript.Echo vbcr

getResource "MSG_INSTALLSUCCESS"
tmpValue = globalValue
getResource "MSG_ENTERKEYPROMPT"

'Prompt to install/activate key
If isCmdline = False Then
    intSuccess = WshShell.Popup (tmpValue & vbCr & vbCr & globalValue,,VALUE_APPNAME, wshYesNoDialog + VALUE_ICON_QUESTIONMARK)
    
    'User declined key entry
    If intSuccess = 7 Then
        getResource "MSG_SLMGR"
        WshShell.Popup globalValue,,VALUE_APPNAME, wshOK + VALUE_ICON_INFORMATION
        pauseExit()
    End If
    
    getResource "MSG_ENTERKEY"
    strProductKey = inputbox(globalValue, VALUE_APPNAME)
End If

If Len(strProductKey) <> 29 Then
    getResource "MSG_UNRECOGKEY"
    tmpValue = globalValue
    getResource "MSG_RERUN"
    WshShell.Popup tmpValue & vbCr & globalValue,,VALUE_APPNAME, wshOK + VALUE_ICON_WARNING
    pauseExit()
End If

'Install Key
Err.Clear()
objSpp.InstallProductKey(strProductKey)
errHandle(folder)

Set winproductinstances = objWMIService.InstancesOf("SoftwareLicensingProduct")
isOffLic = False

'Activate Key
For Each instance in winproductinstances
    instance.refresh_
    If (LCase(instance.ID) = OfficeKmsSkuid) Then
        If instance.ProductKeyID <> "" Or instance.ProductKeyID <> null Then
            isOffLic = True
            instance.Activate
            errHandle(folder)
            getResource "MSG_KEYSUCCESS"
            tmpValue = globalValue
            getResource "MSG_SLMGR"
            If isCmdline = False Then
                WshShell.Popup tmpValue & vbCr & vbCr & globalValue,,VALUE_APPNAME, wshOK + VALUE_ICON_INFORMATION
                pauseExit()
            Else
                WScript.Echo tmpValue & " " & globalValue
                WScript.Quit
            End If
        End If
    End If
Next

If isOffLic = False Then
    getResource "MSG_NOOFFICELIC"
    tmpValue = globalValue
    getResource "MSG_RERUN"
    WshShell.Popup tmpValue & " " & strProductKey & vbCr & vbCr & globalValue,,VALUE_APPNAME, wshOK + VALUE_ICON_WARNING
End If

pauseExit()        
    
End Sub
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function getEngine()

strEngine = LCase(Right(WScript.FullName,12))
If strEngine <> "\cscript.exe" Then
    getResource "MSG_CSCRIPT_NONE"
    WshShell.Popup globalValue & " cscript " & WSCript.ScriptName,,VALUE_APPNAME, wshOK + VALUE_ICON_WARNING
    WScript.Quit
End If

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function CheckSPP(wmiObject)
    Dim colListOfServicesRefresh, objService, installed, running

    installed = False
    running = False
    Set colListOfServicesRefresh = wmiObject.ExecQuery("Select * from Win32_Service ")

    For Each objService in colListOfServicesRefresh
        If objService.Name = "sppsvc" Then
            installed = True
            If LCASE(objService.State) = "running" Then
                running = True
            End If
            Exit For
        End If
    Next

    If installed <> True Then
        getResource "MSG_KMS_ERR_NOINSTALL"
        WshShell.Popup globalValue,,VALUE_APPNAME, wshOK + VALUE_ICON_WARNING
        pauseExit()
    Else
        If running <> True Then
            getResource "MSG_KMS_ERR_NORUN"
            WshShell.Popup globalValue & " " & objService.State,,VALUE_APPNAME, wshOK + VALUE_ICON_WARNING
            pauseExit()
        End If
    End If
    
End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function pauseExit()

getResource "MSG_EXIT"
WScript.Echo globalValue

Do While Not WScript.StdIn.AtEndOfLine
    Input = WScript.StdIn.Read(1)
Loop
    
WScript.Quit

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function getResource(resource)

Set xmlDoc = CreateObject("Msxml2.DOMDocument") 
xmlDoc.load(workingDir & "kms_host.xml")  
Set ElemList = xmlDoc.getElementsByTagName(resource) 
resValue = ElemList.item(0).text
globalValue = resValue 

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function verifyFileExists(file)

If Not fso.FileExists(workingDir & file) Then
    WshShell.Popup MSG_RESOURCE_NONE & vbCr & workingDir & file,,VALUE_APPNAME, wshOK + VALUE_ICON_WARNING
    WScript.Quit
End If

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function errHandle(OS)

Select Case Err.Number
    Case 0
        'Success
    Case Else
        If Hex(Err.Number) = "C004F050" Then
            getResource "MSG_ERRCODE"
            tmpValue = globalValue
            getResource "MSG_SL_E_INVALID_PRODUCT_KEY"
            tmpValue1 = globalValue
            getResource "MSG_RERUN"
            WshShell.Popup tmpValue & " 0x" & Hex(Err.Number) & vbCr & tmpValue1 & vbCr & globalValue,,VALUE_APPNAME, wshOK + VALUE_ICON_WARNING
            pauseExit()
        End If
        
        Select Case OS
            Case "win7", "win8"
                getResource "MSG_ERRCODE"
                tmpValue = globalValue
                getResource "MSG_SLUI"
                WshShell.Popup tmpValue & " 0x" & Hex(Err.Number) & vbCr & globalValue & vbCr & MSG_SLUIEXE & Hex(Err.Number),,VALUE_APPNAME, wshOK + VALUE_ICON_WARNING
            Case Else
                getResource "MSG_ERRCODE"
                WshShell.Popup globalValue & " 0x" & Hex(Err.Number),,VALUE_APPNAME, wshOK + VALUE_ICON_WARNING
        End Select
        pauseExit()
End Select

Err.Clear

End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
Function InstallKMSLicense(objSPP, licFile, OS)

    getResource "MSG_INSTALLLIC"
    WScript.Echo globalValue & " " & licFile

    LicenseData = ReadAllTextFile(licFile)
    errHandle(OS)

    objSpp.InstallLicense(LicenseData)
    errHandle(OS)
    
End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
' Returns the encoding for a givven file.
' Possible return values: ascii, unicode, unicodeFFFE (big-endian), utf-8
Function GetFileEncoding(strFileName)
    Dim strData
    Dim strEncoding

    Set oStream = CreateObject("ADODB.Stream")

    oStream.Type = 1 'adTypeBinary
    oStream.Open
    oStream.LoadFromFile(strFileName)

    ' Default encoding is ascii
    strEncoding =  "ascii"

    strData = BinaryToString(oStream.Read(2))

    ' Check for little endian (x86) unicode preamble
    If (Len(strData) = 2) and strData = (Chr(255) + Chr(254)) Then
        strEncoding = "unicode"
    Else
        oStream.Position = 0
        strData = BinaryToString(oStream.Read(3))

        ' Check for utf-8 preamble
        If (Len(strData) >= 3) and strData = (Chr(239) + Chr(187) + Chr(191)) Then
            strEncoding = "utf-8"
        End If
    End If

    oStream.Close

    GetFileEncoding = strEncoding
    
End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
' Converts binary data (VT_UI1 | VT_ARRAY) to a string (BSTR)
Function BinaryToString(dataBinary)  
    Dim i
    Dim str

    For i = 1 To LenB(dataBinary)
        str = str & Chr(AscB(MidB(dataBinary, i, 1)))
    Next

    BinaryToString = str
    
End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
' Returns string containing the whole text file data. 
' Supports ascii, unicode (little-endian) and utf-8 encoding.
Function ReadAllTextFile(strFileName)
    Dim strData
    Set oStream = CreateObject("ADODB.Stream")

    oStream.Type = 2 'adTypeText
    oStream.Open
    oStream.Charset = GetFileEncoding(strFileName)
    oStream.LoadFromFile(strFileName)

    strData = oStream.ReadText(-1) 'adReadAll

    oStream.Close

    ReadAllTextFile = strData
    
End Function
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////