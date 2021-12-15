On Error Resume Next
Set objService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
If Err.Number <> 0 Then
	WScript.Echo Err.Number & ": " & Err.Description
	WScript.Quit
End If
For Each objSound In objService.ExecQuery("SELECT * FROM Win32_SoundDevice")
	WScript.Echo objSound.Caption 'наименование устройства
	WScript.Echo objSound.ProductName 'наименование устройства
	WScript.Echo objSound.Description 'описание устройства
	WScript.Echo objSound.Manufacturer 'производитель
	WScript.Echo objSound.DeviceID 'идентификатор устройства
	WScript.Echo objSound.SystemName 'имя компьютера
Next
For Each objPnP In objService.ExecQuery("SELECT * FROM Win32_PnPEntity")
	WScript.Echo objPnP.Name 'наименование устройства
	WScript.Echo objPnP.Description 'описание устройства
	WScript.Echo objPnP.Manufacturer 'производитель
	WScript.Echo objPnP.PNPDeviceID 'идентификатор логического устройства
	WScript.Echo
Next