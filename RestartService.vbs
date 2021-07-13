	
	Option Explicit

'Restart services
'Go to services and double click on a service to get the name of the service
'If the computer is local, put a "dot" for computer argument
RestartService ".", "AAE_AutoLoginService_v11"

Sub RestartService(Computer, ServiceName)
	  
		StopService Computer, ServiceName, True
	 
		StartService Computer, ServiceName, True 
	  
End Sub

Sub StopService(Computer, ServiceName, Wait)
	  Dim cimv2, oService, Result

	  'Get the WMI administration object    
	  Set cimv2 = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		Computer & "\root\cimv2")

	  'Get the service object
	  Set oService = cimv2.Get("Win32_Service.Name='" & ServiceName & "'")
	  
	  'Check base properties
	  If Not oService.Started Then
		' the service is Not started
		'wscript.echo "The service " & ServiceName & " is Not started"
		exit Sub
	  End If

	  If Not oService.AcceptStop Then
		' the service does Not accept stop command
		'wscript.echo "The service " & ServiceName & " does Not accept stop command"
		exit Sub
	  End If
	  
	  'wscript.echo oService.getobjecttext_

	  'Stop the service
	  Result  = oService.StopService
	  If 0 <> Result Then
		'wscript.echo "Stop " & ServiceName & " error: " & Result
		exit Sub 
	  End If 
	  
	  Do While oService.Started And Wait
		'get the current service state
		Set oService = cimv2.Get("Win32_Service.Name='" & ServiceName & "'")

		'wscript.echo now, "StopService", ServiceName, oService.Started, oService.State, oService.Status
		Wscript.Sleep 200
	  Loop   
End Sub


Sub StartService(Computer, ServiceName, Wait)
	  Dim cimv2, oService, Result

	  'Get the WMI administration object    
	  Set cimv2 = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		Computer & "\root\cimv2")

	  'Get the service object
	  Set oService = cimv2.Get("Win32_Service.Name='" & ServiceName & "'")

	  'Check base properties
	  If oService.Started Then
		' the service is Not started
		'wscript.echo "The service " & ServiceName & " is started."
		exit Sub
	  End If
	  
	  'Start the service
	  Result = oService.StartService
	  If 0 <> Result Then
		'wscript.echo "Start " & ServiceName & " error:" & Result
		exit Sub 
	  End If 
	  
	  Do While InStr(1,oService.State,"running",1) = 0 And Wait 
		'get the current service state
		Set oService = cimv2.Get("Win32_Service.Name='" & ServiceName & "'")
		
		'wscript.echo now, "StartService", ServiceName, oService.Started, oService.State, oService.Status
		Wscript.Sleep 200
	  Loop   
End Sub
