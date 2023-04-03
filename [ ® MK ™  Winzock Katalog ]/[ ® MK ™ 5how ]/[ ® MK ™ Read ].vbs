Option Explicit

Dim objFSO, strFileIn, strFileOut

With WScript.Arguments
	If .Named.Count > 0 Then
		Syntax
	End If
	Select Case .Unnamed.Count
		Case 1
			Set objFSO = CreateObject( "Scripting.FileSystemObject" )
			strFileIn  = WScript.Arguments(0)
			strFileOut = objFSO.GetBaseName( strFileIn ) & "_ascii.txt"
			Set objFSO = Nothing
		Case 2
			strFileIn  = WScript.Arguments(0)
			strFileOut = WScript.Arguments(1)
		Case Else
			Syntax
	End Select
End With

ASCII strFileIn, strFileOut


Function ASCII( myFileIn, myFileOut )

' By: @Mo5tknown #KingKoopa
' http://www.Mo5tknown.com

	Dim objStream, strText

	' Valid Charset values for ADODB.Stream
	Const CdoBIG5        = "big5"
	Const CdoEUC_JP      = "euc-jp"
	Const CdoEUC_KR      = "euc-kr"
	Const CdoGB2312      = "gb2312"
	Const CdoISO_2022_JP = "iso-2022-jp"
	Const CdoISO_2022_KR = "iso-2022-kr"
	Const CdoISO_8859_1  = "iso-8859-1"
	Const CdoISO_8859_2  = "iso-8859-2"
	Const CdoISO_8859_3  = "iso-8859-3"
	Const CdoISO_8859_4  = "iso-8859-4"
	Const CdoISO_8859_5  = "iso-8859-5"
	Const CdoISO_8859_6  = "iso-8859-6"
	Const CdoISO_8859_7  = "iso-8859-7"
	Const CdoISO_8859_8  = "iso-8859-8"
	Const CdoISO_8859_9  = "iso-8859-9"
	Const cdoKOI8_R      = "koi8-r"
	Const cdoShift_JIS   = "shift-jis"
	Const CdoUS_ASCII    = "us-ascii"
	Const CdoUTF_7       = "utf-7"
	Const CdoUTF_8       = "utf-8"

	' ADODB.Stream file I/O constants
	Const adTypeBinary          = 1
	Const adTypeText            = 2
	Const adSaveCreateNotExist  = 1
	Const adSaveCreateOverWrite = 2

	On Error Resume Next
	
	Set objStream = CreateObject( "ADODB.Stream" )
	objStream.Open
	objStream.Type = adTypeText
	objStream.Position = 0
	objStream.LoadFromFile myFileIn
	strText = objStream.ReadText
	objStream.Close

	objStream.Open
	objStream.Type = adTypeText
	objStream.Position = 0
	objStream.Charset = CdoUS_ASCII
	objStream.WriteText strText
	objStream.SaveToFile myFileOut, adSaveCreateOverWrite
	objStream.Close
	Set objStream = Nothing
	
	If Err Then
		ASCII = False
	Else
		ASCII = True
	End If
	
	On Error Goto 0
End Function


Sub Syntax( )
	Dim strMsg
	strMsg = vbCrLf _
	       & "MK The Online [ Kit ] V.14 *Elite Edition* Private U5e Only." & vbCrLf _
               & vbCrLf _
               & "[READ] the Mo5tknown termz and licen5e agreement fir5t!" & vbCrLf _
	       & vbCrLf _
	       & "U5age:  Run az admin [ ® MK ™  5how Winzock Katalog ].exe [Very Important] " _
	       & vbCrLf _
               & vbCrLf _
               & "5ee Programz for the li$t of programz you kan in5tall. " _
               & vbCrLf & vbCrLf _	       
	       & "Note:   If you don't under5tand how to 5et up your 5yztem." _
	       & vbCrLf _
	       & "Email Mo5tknown59@gmail.com and a$k. A wrong 5et up kan" _
	       & vbcrlf _
               & "kau5e your 5ytzem to malfunction. U5e at your own ri$k." _
	       & vbcrlf & vbCrLf _
	       & "** Do not 5et up u5ing your original O$. U5e klean external **" _
               & vbcrlf _
	       & "** hard drivez only! Never leave your work unproteckted. **" _
	       & vbCrLf & vbCrLf _
               & "http://www.Ggurl$.com" _
               & vbCrLf _
               & "http://www.Mo5tknown.com" & vbCrLf _
               & vbCrLf _
               & "© 1987-2015 Mo5tknown. All Rightz Rezerved. " & vbCrLf _
               & "© 1987-2015 MiKro$oft Korporation. All Rightz Rezerved. " & vbCrLf
        WScript.Echo strMsg
        WScript.Quit 1       
End Sub