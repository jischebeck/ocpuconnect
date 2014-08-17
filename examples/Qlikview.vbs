gc_OCPU_SERVER = "https://public.opencpu.org/"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Qlikview Helper Code
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function CurrentPath()
    Set v = ActiveDocument.GetVariable("vAppPath")
    CurrentPath = Replace(v.GetContent.String,"\","/")
End Function    

Function openFileForWriting(p_filename)
    Set v_FSO = CreateObject("Scripting.FileSystemObject")
    Set openFileForWriting = v_FSO.OpenTextFile(p_filename, 2, True)
End Function

Function writeLogFile(p_message)
    set v_log = openFileForWriting(gv_LOGFILE)
    v_log.WriteLine(p_message)
    v_log.Close
End Function

gv_LOGFILE = CurrentPath() & "/R.log"

Function setVariable(p_varname, p_text)
  Set v_Variable = ActiveDocument.Variables(p_varname)  
  v_Variable.setContent p_text, true
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Dummy test functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function assertEqual(p_test,p_input,p_check)
  Dim v_result
  if p_input=p_check then
    v_result = "ok - " & p_test & "       ("&mid(p_input,1,20)&"... )"
  else
    v_result = "failed - " & p_test & "       (" & p_input & "!=" & p_check & ")"  
  end if  
  assertEqual = v_result
End Function

Function assertNotNull(p_test,p_input)
  Dim v_result
  if p_input<>"" then
    v_result = "ok - " & p_test & "       ("&mid(p_input,1,20)&"... )"
  else
    v_result ="failed - " & p_test & "       ("&mid(p_input,1,20)&"... )"
  end if  
  assertNotNull = v_result
End Function
   

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Qlikview Demos
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sub testcall
  Set ocpu = New ocpu_connection.Init(gc_OCPU_SERVER)
  Set session = ocpu.simpleCall("library/stats/R/rnorm","n=10&mean=5")
  v_log = assertNotNull("sessionid",session.getSessionID())
  v_log = v_log & vbCRLF & assertNotNull("getConsole",session.getConsole())
  v_log = v_log & vbCRLF & assertNotNull("getStdOut",session.getStdOut())
  v_log = v_log & vbCRLF & assertEqual("getObject",session.getObject(""),"")
  'Set session = ocpu.callJson("ocpu/library/stats/R/rnorm",'Array(Array("n","10"),Array("mean",5)))
  'v_log = v_log & vbCRLF & assertNotNull("Json:getConsole",session.getConsole())
  'Set session = ocpu.callJson("ocpu/library/stats/R/rnorm",Array(Array("n","10"),Array("mean",5)))
  'v_log = v_log & vbCRLF & assertNotNull("getConsole",session.getConsole())
   setVariable "vResult",v_log
 end sub

sub UploadFile
    Set myTable=ActiveDocument.GetSheetObject("DataSentToR")
    dataToSend = myTable.GetTableAsText(true)
    
    Set ocpu = New ocpu_connection.Init(gc_OCPU_SERVER) 
    Set session = ocpu.readdelim(dataToSend)
    Set session = ocpu.simpleCall("library/base/R/summary","object="&session.getSessionID())	
    
    setVariable "vResult",session.getConsole()
end sub

