 '*********************************************************
 ' OCPU Connect
 '
 ' Version 0.1
 '
 ' This VBScipt Library allows to connect to an OpenCPU Server
 ' (see: https://www.opencpu.org/) and execute R scripts and
 ' functions using the REST API.
 ' It requires the Microsoft.XMLHTTP Objekt for RPC Calls.
 '
 ' Source: https://github.com/jischebeck/ocpuconnect
 '*********************************************************
 
Class ocpu_call
   Private cv_server
   Private cv_command_path
   Private cv_files
   Private cv_params
   Public Default Function Init(p_server,p_command_path)
      Set cv_server = p_server
	  cv_command_path = p_command_path
      Set cv_files = CreateObject("Scripting.Dictionary")
      Set cv_params = CreateObject("Scripting.Dictionary")
	  Set Init = Me
   End Function
   Public Function AddFile(p_name, p_content)
      cv_files.add p_name, p_content
	  Set AddFile = Me
   End Function
   Public Function AddStrParam(p_name, p_content)
      cv_params.add p_name, p_content
	  Set AddStrParam = Me
   End Function
   Public Function exec
      Set exec = cv_server.complexCall(cv_command_path,cv_params,cv_files)
   End Function   
End Class
	
Class ocpu_session
   Private cv_server
   Private cv_session_id
   Private cv_location
   Public Default Function Init(p_server,p_session_id,p_location)
      Set cv_server = p_server
	  cv_session_id = p_session_id
	  cv_location = p_location
	  Set Init = Me
   End Function

   Public Function getSessionID
	  getSessionID = cv_session_id
   End Function   
   
   Public Function getLoc
	  getLoc = cv_location
   End Function

   Function getObject(p_name)
      if p_name="" then
	     v_name = ".val"
	  else
	     v_name = p_name
	  end if
	  getObject = cv_server.httpGet("/ocpu/tmp/" & cv_session_id & "/R/" & v_name & "/json")
	  'Set html = CreateObject("htmlfile")
      'Set window = html.parentWindow
      'window.execScript "var json = " & strJson, "JScript"
      'v_result = window.json
   End Function
   
   Function getFileURL(p_path)
      getFileURL = getLoc() & "files/" & p_path
   End Function
   
   Function getFile(p_path)
      getFile = cv_server.httpGet(getFileURL(p_path))
   End Function

   ' Retrieve plot number {n} in output format {format}. Format is usually one of png, pdf or svg. The {n} is an integer or "last".   
   Function getGraphicsURL(p_nr,p_format)    
      getGraphicURL = getLoc() & "graphics/" & p_nr & "/" & p_format 
   End Function
   
   Function getGraphics(p_path)
      getGraphic = cv_server.httpGet(getGraphicsURL(p_nr,p_format))
   End Function
   
   Function getConsole
	  getConsole = cv_server.httpGet("/ocpu/tmp/" & cv_session_id & "/console/text")
   End Function
   
   Function getSource
	  getSource = cv_server.httpGet("/ocpu/tmp/" & cv_session_id & "/source/text")
   End Function

   Function getStdout
	  getStdout = cv_server.httpGet("/ocpu/tmp/" & cv_session_id & "/stdout/text")
   End Function
   
   Function exec(p_command,p_param)
	  Set exec = cv_server.simpleCall(getLoc() & "R/" & p_command, p_param)
   End Function
End Class

Class ocpu_connection
   Public cv_server
   Public Default Function Init(p_server)
      cv_server = p_server
      Set Init = Me
   End Function
   
   Public Function simpleCall(p_command,p_param)
      ' TODO: add parameter checking
	  ' is p_param Array, Dictionary or ...
	  v_url = cv_server & "ocpu/" & p_command
	  Set v_xmlhttp = CreateObject("Microsoft.XMLHTTP")
	  v_xmlhttp.open "POST", v_url, false 
	  v_xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	  v_xmlhttp.send p_param
      v_result = v_xmlhttp.responseText
	  v_xmlhttp = null
	  v_session_id = mid(v_result,11, 11)
      ' TODO: Switch to getResponseHeader('Location')  and getResponseHeader('X-ocpu-session')	  
      ' TODO: Add Error Handling
	  Set simpleCall = New ocpu_session.Init(Me,v_session_id,v_result)
   End Function

   Public Function httpGet(p_path)
      v_url = cv_server & p_path
	  Set v_xmlhttp = CreateObject("Microsoft.XMLHTTP")
	  v_xmlhttp.open "GET", v_url, false
      v_xmlhttp.setRequestHeader "Content-Type", "text/xml"
	  v_xmlhttp.send ""
      httpGet = v_xmlhttp.responseText	
	  ' TODO: Add Error Handling
	  v_xmlhttp = Null
   End Function
   
   Public Function createCall(p_command)
     Set createCall = New ocpu_call.Init(Me, p_command)
   End Function
   
   Public Function readdelim(p_content)
      Set readdelim = createCall("library/utils/R/read.delim/").addFile("file",p_content).exec()
   End Function
   
   Public Function complexCall(p_path,p_params,p_files)
      v_url = cv_server & "ocpu/" & p_path
	  v_param_names = p_params.Keys
	  v_param_values = p_params.Items
	  For i = 0 to p_params.Count-1
	     v_params = v_params & v_param_names(i) & "=" & v_param_values (i) & "&"
	  Next
	  Set v_xmlhttp = CreateObject("Microsoft.XMLHTTP")
      v_xmlhttp.open "POST", v_url & "?" & v_params, false 
		  ' multipart handling derived from http://stackoverflow.com/questions/5933949/how-to-send-multipart-form-data-form-content-by-ajax-no-jquery
	  v_BOUNDARY = "3fbd04f5-b1ed-4060-99b9-fca7ff59c113"
	  v_xmlhttp.setRequestHeader "Content-Type", "multipart/form-data; charset=utf-8; boundary=" + v_BOUNDARY
	  
	  ' Build POST for Multiparts
      v_file_contents = p_files.Items 
	  v_file_param_names = p_files.Keys
      v_multipart = "--" & v_BOUNDARY 
	  For i = 0 To p_files.Count -1 
	     v_multipart = v_multipart & vbCrLf & _
        "Content-Disposition: form-data; name="""&v_file_param_names(i)&"""; filename=""file"& i & """" & vbCrLf & _
        "Content-Type: application/octet-stream" & vbCrLf & vbCrLf & _
         v_file_contents(i) & vbCrLf & "--" & v_BOUNDARY & "--"
	  Next  
      v_multipart = v_multipart & "--"
	  v_xmlhttp.send v_multipart
	
      v_result = v_xmlhttp.responseText	
	  v_session_id = mid(v_result,11, 11)
      ' TODO: Switch to getResponseHeader('Location')  and getResponseHeader('X-ocpu-session')	  
	  ' TODO: Add Error Handling
	  Set complexCall = New ocpu_session.Init(Me,v_session_id,v_result)
   End Function
 End Class  
  