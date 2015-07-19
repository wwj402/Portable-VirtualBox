#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <GUIConstantsEx.au3>
#include <ButtonConstants.au3>
#include <File.au3>


Local $HOME_BASE, $arch
$HOME_BASE = _PathFull(@ScriptDir &"\..\..", "\")
;~ MsgBox (0, "test", $HOME_BASE)

If FileExists ($HOME_BASE&"\app32\") AND FileExists ($HOME_BASE&"\app64\") Then
	If @OSArch = "x86" Then
      Global $arch = "app32"
    EndIf
    If @OSArch = "x64" Then
      Global $arch = "app64"
    EndIf
Else
    If FileExists ($HOME_BASE&"\app32\") AND NOT FileExists ($HOME_BASE&"\app64\") Then
      Global $arch = "app32"
    EndIf
    If NOT FileExists ($HOME_BASE&"\app32\") AND FileExists ($HOME_BASE&"\app64\") Then
      Global $arch = "app64"
    EndIf
EndIf

If FileExists ($HOME_BASE&"\Portable-VirtualBox.exe") Then
	MsgBox (0, "HOME FOLDER", "Root folder is " & $HOME_BASE & ".")
Else
	Local $PathHR = FileSelectFolder ("Select root folder", $HOME_BASE)
	If NOT @error Then
		$HOME_BASE =  $PathHR
		MsgBox (0, "HOME FOLDER", "Root folder is " & $HOME_BASE & ".")
	EndIf
EndIf


Opt('MustDeclareVars', 1)

DriverGUI()

Func DriverGUI()
    Local $Button_1, $Button_2, $Button_3, $Button_4, $Button_5, $msg
    GUICreate("Manual install/remove drivers", 400, 200) ; will create a dialog box that when displayed is centered

    Opt("GUICoordMode", 2)
    $Button_1 = GUICtrlCreateButton("Install the drivers", 20, 25, 180, 40, $BS_CENTER + $BS_VCENTER + $BS_MULTILINE)
	$Button_2 = GUICtrlCreateButton("Remove the drivers", 4, -1, 180, 40, $BS_CENTER + $BS_VCENTER + $BS_MULTILINE)
	$Button_3 = GUICtrlCreateButton("Install the servics", -364, 15, 180, 40, $BS_CENTER + $BS_VCENTER + $BS_MULTILINE)
	$Button_4 = GUICtrlCreateButton("Remove the services", 4, -1, 180, 40, $BS_CENTER + $BS_VCENTER + $BS_MULTILINE)
	$Button_5 = GUICtrlCreateButton("Run portable virtualbox", -1, 15, 180, 40, $BS_CENTER + $BS_VCENTER + $BS_MULTILINE)

    GUISetState()      ; will display an  dialog box with 2 button

    ; Run the GUI until the dialog is closed
    While 1
        $msg = GUIGetMsg()
        Select
            Case $msg = $GUI_EVENT_CLOSE
                ExitLoop
            Case $msg = $Button_1
                InstallDrivers()   ; Will install drivers
            Case $msg = $Button_2
                RemoveDrivers()    ; Will remove drivers
            Case $msg = $Button_3
				InstallServices()
            Case $msg = $Button_4
				RemoveServices()
			Case $msg = $Button_5
				If ProcessExists("VirtualBox.exe") Then
					MsgBox(0, "Portable-VirtualBox", "VirtualBox is running. No need run Portable-VirtualBox.")
				Else
					Run($HOME_BASE &"\Portable-VirtualBox.exe")
				EndIf

        EndSelect
    WEnd
EndFunc   ;==>


Func InstallDrivers()
	Local $REG_VALUE
	If ProcessExists("VirtualBox.exe") Then
		MsgBox(0, "Install Drivers", "VirtualBox is running. Please close virtalbox to install drivers.")
	Else
;~ 		MsgBox (0, "test", RegRead ("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\VBoxUSB", "DisplayName"))
		If RegRead ("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\VBoxUSB", "DisplayName") <> "VirtualBox USB" Then
			InstallUsbDevice()
;~ 			MsgBox (0, "test", "Install UsbDevice")
		Else
			MsgBox (0, "VBoxUSB", "Device VBoxUSB is installed.")
;~ 			RunWait ("sc start VBoxUSB", $HOME_BASE, @SW_HIDE)
		EndIf

;~ 		MsgBox (0, "test", RegRead ("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\VBoxNetAdp", "DisplayName"))
		If RegRead ("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\VBoxNetAdp", "DisplayName") <> "VirtualBox Host-Only Ethernet Adapter" Then
			InstallNetAdp()
;~ 			MsgBox (0, "test", "Install netadp")
		Else
			MsgBox (0, "VBoxNetAdp", "Device VBoxNetAdp is installed.")
;~ 			RunWait ("sc start VBoxNetAdp", $HOME_BASE, @SW_HIDE)
		EndIf

;~ 		MsgBox (0, "test", RegRead ("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\VBoxNetFlt", "DisplayName"))
		If RegRead ("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\VBoxNetFlt", "DisplayName") <> "VirtualBox Bridged Networking Service" Then
			InstallNetFlt()
;~ 			MsgBox (0, "test", "Install netflt")
		Else
			MsgBox (0, "VBoxNetFlt", "Device VBoxNetFlt is installed.")
;~ 			RunWait ("sc start VBoxNetFlt", $HOME_BASE, @SW_HIDE)
		EndIf
	EndIf
EndFunc

Func RemoveDrivers()
	If ProcessExists("VirtualBox.exe") Then
		MsgBox(0, "Remove Drivers", "VirtualBox is running. Please close virtalbox to remove drivers.")
	Else
		RemoveUsbDevice()
		RemoveNetAdp()
		RemoveNetFlt()
	EndIf
EndFunc


Func InstallServices()
	If ProcessExists("VirtualBox.exe") Then
		MsgBox(0, "Install services", "VirtualBox is running. Please close virtalbox to install services.")
	Else
;~ 		MsgBox (0, "testreg", RegRead ("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\VBoxUSBMon", "ImagePath"))
;~ 		MsgBox (0, "testpath", "\??\" & $HOME_BASE &"\"& $arch &"\drivers\USB\filter\VBoxUSBMon.sys")
;~ 		If RegRead ("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\VBoxUSBMon", "DisplayName") <> "VirtualBox USB Monitor Driver" Then
		If RegRead ("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\VBoxUSBMon", "ImagePath") <> "\??\" &$HOME_BASE &"\"& $arch &"\drivers\USB\filter\VBoxUSBMon.sys" Then
			InstallUsbService()
;~ 			MsgBox (0, "test", "Install Usbmon")
		Else
			MsgBox (0, "VBoxUSBMon", "Service VBoxUSBMon is intalled.")
;~ 			RunWait ("sc start VBoxUSBMon", $HOME_BASE, @SW_HIDE)
		EndIf

;~ 		MsgBox (0, "testreg", RegRead ("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\VBoxDRV", "ImagePath"))
;~ 		MsgBox (0, "testpath", "\??\" & $HOME_BASE &"\"& $arch &"\drivers\vboxdrv\VBoxDrv.sys")
		If RegRead ("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\VBoxDRV", "ImagePath") <> "\??\" & $HOME_BASE &"\"& $arch &"\drivers\vboxdrv\VBoxDrv.sys" Then
;~ 		If RegRead ("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\VBoxDRV", "DisplayName") <> "VirtualBox Service" Then
			InstallVBoxDRV()
;~ 			MsgBox (0, "test", "install VBoxDRV")
		Else
			MsgBox (0, "VBoxDRV", "Service VBoxDRV is installed.")
;~ 			RunWait ("sc start VBoxDRV", $HOME_BASE, @SW_HIDE)
		EndIf
	EndIf
EndFunc


Func RemoveServices()
	If ProcessExists("VirtualBox.exe") Then
		MsgBox(0, "Remove Drivers", "VirtualBox is running. Please close virtalbox to remove drivers.")
	Else
		RemoveServicefunc("VBoxDRV")
		RemoveServicefunc("VBoxUSBMon")
	EndIf
EndFunc



Func InstallUsbDevice()
	Local $ERRORCODE1
	If @OSArch = "x86" Then
		$ERRORCODE1 = RunWait ($HOME_BASE &"\data\tools\devcon_x86.exe install .\"& $arch &"\drivers\USB\device\VBoxUSB.inf ""USB\VID_80EE&PID_CAFE""", $HOME_BASE, @SW_HIDE)
	EndIf
	If @OSArch = "x64" Then
		$ERRORCODE1 = RunWait ($HOME_BASE &"\data\tools\devcon_x64.exe install .\"& $arch &"\drivers\USB\device\VBoxUSB.inf ""USB\VID_80EE&PID_CAFE""", $HOME_BASE, @SW_HIDE)
	EndIf
	Local $ERRORCODE2
	FileCopy ($HOME_BASE&"\"& $arch &"\drivers\USB\device\VBoxUSB.sys", @SystemDir&"\drivers", 9)
	$ERRORCODE2 = RunWait ("sc start VBoxUSB", $HOME_BASE, @SW_HIDE)
	MsgBox (0, "Error code", "Install driver code: "& $ERRORCODE1 &"; Start service code: "& $ERRORCODE2 &".")
EndFunc

Func InstallNetAdp()
	Local $ERRORCODE1
	If @OSArch = "x86" Then
		$ERRORCODE1 = RunWait ($HOME_BASE &"\data\tools\devcon_x86.exe install .\"& $arch &"\drivers\network\netadp\VBoxNetAdp.inf ""sun_VBoxNetAdp""", $HOME_BASE, @SW_HIDE)
	EndIf
	If @OSArch = "x64" Then
		$ERRORCODE1 = RunWait ($HOME_BASE &"\data\tools\devcon_x64.exe install .\"& $arch &"\drivers\network\netadp\VBoxNetAdp.inf ""sun_VBoxNetAdp""", $HOME_BASE, @SW_HIDE)
	EndIf
	Local $ERRORCODE2
	FileCopy ($HOME_BASE&"\"& $arch &"\drivers\network\netadp\VBoxNetAdp.sys", @SystemDir&"\drivers", 9)
	$ERRORCODE2 = RunWait ("sc start VBoxNetAdp", $HOME_BASE, @SW_HIDE)
	MsgBox (0, "Error code", "Install driver code: "& $ERRORCODE1 &"; Start service code: "& $ERRORCODE2 &".")
EndFunc

Func InstallNetFlt()
	Local $ERRORCODE1
	If @OSArch = "x86" Then
		$ERRORCODE1 = RunWait ($HOME_BASE&"\data\tools\snetcfg_x86.exe -v -u sun_VBoxNetFlt", $HOME_BASE, @SW_HIDE)
		$ERRORCODE1 = RunWait ($HOME_BASE&"\data\tools\snetcfg_x86.exe -v -l .\"& $arch &"\drivers\network\netflt\VBoxNetFlt.inf -m .\"& $arch &"\drivers\network\netflt\VBoxNetFltM.inf -c s -i sun_VBoxNetFlt", $HOME_BASE, @SW_HIDE)
	EndIf
	If @OSArch = "x64" Then
		$ERRORCODE1 = RunWait ($HOME_BASE&"\data\tools\snetcfg_x64.exe -v -u sun_VBoxNetFlt", $HOME_BASE, @SW_HIDE)
		$ERRORCODE1 = RunWait ($HOME_BASE&"\data\tools\snetcfg_x64.exe -v -l .\"& $arch &"\drivers\network\netflt\VBoxNetFlt.inf -m .\"& $arch &"\drivers\network\netflt\VBoxNetFltM.inf -c s -i sun_VBoxNetFlt", $HOME_BASE, @SW_HIDE)
	EndIf
	Local $ERRORCODE2
	FileCopy ($HOME_BASE&"\"& $arch &"\drivers\network\netflt\VBoxNetFltNobj.dll", @SystemDir, 9)
	FileCopy ($HOME_BASE&"\"& $arch &"\drivers\network\netflt\VBoxNetFlt.sys", @SystemDir&"\drivers", 9)
	RunWait (@SystemDir&"\regsvr32.exe /S "& @SystemDir &"\VBoxNetFltNobj.dll", $HOME_BASE, @SW_HIDE)
	$ERRORCODE2 = RunWait ("sc start VBoxNetFlt", $HOME_BASE, @SW_HIDE)
	MsgBox (0, "Error code", "Install driver code: "& $ERRORCODE1 &"; Start service code: "& $ERRORCODE2 &".")
EndFunc



Func RemoveUsbDevice()
	Local $ERRORCODE1
	RunWait ("sc stop VBoxUSB", $HOME_BASE, @SW_HIDE)
	If @OSArch = "x86" Then
		$ERRORCODE1 = RunWait ($HOME_BASE &"\data\tools\devcon_x86.exe remove ""USB\VID_80EE&PID_CAFE""", $HOME_BASE, @SW_HIDE)
	EndIf
	If @OSArch = "x64" Then
		$ERRORCODE1 = RunWait ($HOME_BASE &"\data\tools\devcon_x64.exe remove ""USB\VID_80EE&PID_CAFE""", $HOME_BASE, @SW_HIDE)
	EndIf
	Local $ERRORCODE2
	$ERRORCODE2 = RunWait ("sc delete VBoxUSB", $HOME_BASE, @SW_HIDE)
	FileDelete (@SystemDir&"\drivers\VBoxUSB.sys")
	MsgBox (0, "Error code", "Remove driver code: "& $ERRORCODE1 &"; Delete service code: "& $ERRORCODE2 &".")
EndFunc

Func RemoveNetAdp()
	Local $ERRORCODE1
	RunWait ("sc stop VBoxNetAdp", $HOME_BASE, @SW_HIDE)
	If @OSArch = "x86" Then
		$ERRORCODE1 = RunWait ($HOME_BASE &"\data\tools\devcon_x86.exe remove ""sun_VBoxNetAdp""", $HOME_BASE, @SW_HIDE)
	EndIf
	If @OSArch = "x64" Then
		$ERRORCODE1 = RunWait ($HOME_BASE &"\data\tools\devcon_x64.exe remove ""sun_VBoxNetAdp""", $HOME_BASE, @SW_HIDE)
	EndIf
	Local $ERRORCODE2
	$ERRORCODE2 = RunWait ("sc delete VBoxNetAdp", $HOME_BASE, @SW_HIDE)
	FileDelete (@SystemDir&"\drivers\VBoxNetAdp.sys")
	MsgBox (0, "Error code", "Remove driver code: "& $ERRORCODE1 &"; Delete service code: "& $ERRORCODE2 &".")
EndFunc

Func RemoveNetFlt()
	Local $ERRORCODE1
	RunWait ("sc stop VBoxNetFlt", $HOME_BASE, @SW_HIDE)
	If @OSArch = "x86" Then
		$ERRORCODE1 = RunWait ($HOME_BASE&"\data\tools\snetcfg_x86.exe -v -u sun_VBoxNetFlt", $HOME_BASE, @SW_HIDE)
	EndIf
	If @OSArch = "x64" Then
		$ERRORCODE1 = RunWait ($HOME_BASE&"\data\tools\snetcfg_x64.exe -v -u sun_VBoxNetFlt", $HOME_BASE, @SW_HIDE)
	EndIf
	Local $ERRORCODE2
	RunWait (@SystemDir&"\regsvr32.exe /S /U "&@SystemDir&"\VBoxNetFltNobj.dll", $HOME_BASE, @SW_HIDE)
	$ERRORCODE2 = RunWait ("sc delete VBoxNetFlt", $HOME_BASE, @SW_HIDE)
	FileDelete (@SystemDir&"\VBoxNetFltNobj.dll")
	FileDelete (@SystemDir&"\drivers\VBoxNetFlt.sys")
	MsgBox (0, "Error code", "Remove driver code: "& $ERRORCODE1 &"; Delete service code: "& $ERRORCODE2 &".")
EndFunc



Func InstallUsbService()
	Local $ERRORCODE1
	$ERRORCODE1 = RunWait ("cmd /c sc create VBoxUSBMon binpath= ""%CD%\"& $arch &"\drivers\USB\filter\VBoxUSBMon.sys"" type= kernel start= auto error= normal displayname= PortableVBoxUSBMon", $HOME_BASE, @SW_HIDE)
	Local $ERRORCODE2
	$ERRORCODE2 = RunWait ("sc start VBoxUSBMon", $HOME_BASE, @SW_HIDE)
	MsgBox (0, "Error code", "Create service code: "& $ERRORCODE1 &"; Start service code: "& $ERRORCODE2 &".")
EndFunc

Func InstallVBoxDRV()
	Local $ERRORCODE1
	$ERRORCODE1 = RunWait ("cmd /c sc create VBoxDRV binpath= ""%CD%\"& $arch &"\drivers\vboxdrv\VBoxDrv.sys"" type= kernel start= auto error= normal displayname= PortableVBoxDRV", $HOME_BASE, @SW_HIDE)
	Local $ERRORCODE2
	$ERRORCODE2 = RunWait ("sc start VBoxDRV", $HOME_BASE, @SW_HIDE)
	MsgBox (0, "Error code", "Create service code: "& $ERRORCODE1 &"; Start service code: "& $ERRORCODE2 &".")
EndFunc



Func RemoveServicefunc($SERVICENAME)
	Local $ERRORCODE1
	$ERRORCODE1 = RunWait ("sc stop " & $SERVICENAME, $HOME_BASE, @SW_HIDE)
	Local $ERRORCODE2
	$ERRORCODE2 = RunWait ("sc delete " & $SERVICENAME, $HOME_BASE, @SW_HIDE)
	MsgBox (0, "Error code", "Stop service code: "& $ERRORCODE1 &"; Delete service code: "& $ERRORCODE2 &".")
EndFunc

;~ Func RemoveVBoxDRV()
;~ 	RunWait ("sc stop VBoxDRV", $HOME_BASE, @SW_HIDE)
;~ 	RunWait ("sc delete VBoxDRV", $HOME_BASE, @SW_HIDE)
;~ EndFunc

;~ Func RemoveVBoxUSBMon()
;~ 	RunWait ("sc stop VBoxUSBMon", $HOME_BASE, @SW_HIDE)
;~ 	RunWait ("sc delete VBoxUSBMon", $HOME_BASE, @SW_HIDE)
;~ EndFunc