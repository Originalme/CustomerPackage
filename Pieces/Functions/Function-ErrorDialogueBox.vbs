'---------------------------------------------------------'
'  Function: Displays error message with sMsg as content  '
'---------------------------------------------------------'
Function ErrorBox(sMsg)

	WriteLog("!ERROR: " & sMsg)
	ErrorBox = MsgBox ( _
	"ERROR: " & vbNewLine &_
	sMsg & vbNewLine & vbNewLine &_
	"-------------------------------------------------------------" & vbNewLine & vbNewLine &_
	"Please Contact Forward Advantage." & vbNewLine &_
	"Phone: " & sSupportNumber & vbNewLine &_
	"Email: " & sSupportEmail, _
	vbCritical,sAppName _
	)
End Function