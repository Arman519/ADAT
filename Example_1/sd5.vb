Option Strict Off
'Imports all of the stuff we're going to need
Imports System.DirectoryServices
Imports System.IO
Imports System.Text
Imports System.Collections.ObjectModel
Imports System.Management.Automation
Imports System.Management.Automation.Runspaces

'Creates the Form
Public Class Form1

    'Function to run powershell script
    Private Function RunScript(ByVal scriptText As String) As Object
        Dim MyRunSpace As Runspace = RunspaceFactory.CreateRunspace()
        MyRunSpace.Open()
        Dim MyPipeline As Pipeline = MyRunSpace.CreatePipeline()
        MyPipeline.Commands.AddScript(scriptText)
        Dim results As Collection(Of PSObject) = MyPipeline.Invoke()
        MyRunSpace.Close()
        Dim MyStringBuilder As New StringBuilder()
        For Each obj As PSObject In results
            MyStringBuilder.AppendLine(obj.ToString())
        Next
        Return MyStringBuilder
    End Function
    'load the form
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        TbUserID.Clear()
        SelectedComputer.Clear()
        TbEmail.Clear()
        TbHome.Clear()
        TbMobile.Clear()
        TbIPPhone.Clear()
        TbAddress.Clear()
        TbCity.Clear()
        TbState.Clear()
        TbZip.Clear()
        rtbNotes.Clear()
        rtbDescription.Clear()
        TbUserID.Focus()
    End Sub

    'Gets Selected Computer info Powershell script
    Private Function GetUserComputerSelected()
        Dim Script As New StringBuilder()
        Script.Append("$id = " + Chr(34) + SelectedComputer.Text + Chr(34) + vbCrLf)
        Script.Append("$ComputerName = $id" + vbCrLf)
        Script.Append(" $os = Get-WmiObject win32_operatingsystem -ComputerName $ComputerName -ErrorAction SilentlyContinue" + vbCrLf)
        Script.Append(" $computer = Get-WmiObject win32_computersystem -ComputerName $ComputerName -ErrorAction SilentlyContinue" + vbCrLf)
        Script.Append(" $battery = Get-WmiObject Win32_Battery -ComputerName $ComputerName -ErrorAction SilentlyContinue" + vbCrLf)
        Script.Append(" $bios = Get-WmiObject Win32_Bios -ComputerName $ComputerName -ErrorAction SilentlyContinue" + vbCrLf)
        Script.Append(" $network = gwmi Win32_NetworkAdapterConfiguration  -ComputerName $ComputerName -ErrorAction SilentlyContinue | Where {$_.IPAddress} " + vbCrLf)
        Script.Append(" $cpu = Get-WmiObject win32_processor -computername $computername" + vbCrLf)
        Script.Append(" $mem = Get-WmiObject win32_operatingsystem -ComputerName $ComputerName | Foreach {" + Chr(34) + "{0:N2}" + Chr(34) + " -f ((($_.TotalVisibleMemorySize - $_.FreePhysicalMemory)*100)/ $_.TotalVisibleMemorySize)}" + vbCrLf)
        Script.Append(" $theuser = (get-wmiobject Win32_ComputerSystem -ComputerName $ComputerName).UserName.Split('\')[1]" + vbCrLf)
        Script.Append(" $cpuuseaverage = $cpu | Measure-Object -property LoadPercentage -Average" + vbCrLf)
        Script.Append(" $thecomputerserial = $bios.SerialNumber" + vbCrLf)
        Script.Append(" $uptime = (Get-Date) - $os.ConvertToDateTime($os.LastBootUpTime)" + vbCrLf)
        Script.Append(" $thecomputername = $computer.Name" + vbCrLf)
        Script.Append(" $thecomputermodel = $computer.Model" + vbCrLf)
        Script.Append(" $thecomputermake = $computer.Manufacturer" + vbCrLf)
        Script.Append(" $thebiosversion = $bios.SMBIOSBIOSVersion" + vbCrLf)
        Script.Append(" $theipaddress = $network.ipaddress.Item(0)" + vbCrLf)
        Script.Append(" $thecpuname = $cpu.Name" + vbCrLf)
        Script.Append(" $thecpuloadaverage = $cpuuseaverage.Average" + vbCrLf)
        Script.Append(" $theos = $OS.Caption + " + Chr(34) + " " + Chr(34) + " + $os.OSArchitecture" + vbCrLf)
        Script.Append(" $InstalledRAM = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $ComputerName" + vbCrLf)
        Script.Append(" $installedmemory = [Math]::Round(($InstalledRAM.TotalPhysicalMemory/ 1GB),2)" + vbCrLf)
        Script.Append("" + vbCrLf)
        Script.Append("$dsk = Get-WmiObject Win32_LogicalDisk -ComputerName $Env:ComputerName -ErrorAction SilentlyContinue | Select DeviceID,@{Name=" + Chr(34) + "TotalGB" + Chr(34) + ";Expression={$_.Size/1GB -as [int]}},@{Name=" + Chr(34) + "FreeGB" + Chr(34) + ";Expression={[math]::Round($_.Freespace/1GB,2)}}" + vbCrLf)
        Script.Append("$disk0 = $dsk.DeviceID.Item(0)" + vbCrLf)
        Script.Append("$disk0free = $dsk.FreeGB.Item(0)" + vbCrLf)
        Script.Append("$disk0total = $dsk.TotalGB.Item(0)" + vbCrLf)
        Script.Append("" + vbCrLf)
        Script.Append("" + vbCrLf)
        Script.Append("$hotfixes = " + Chr(34) + "KB4012212" + Chr(34) + "," + Chr(34) + "KB4012213" + Chr(34) + "," + Chr(34) + "KB4012214" + Chr(34) + "," + Chr(34) + "KB4012215" + Chr(34) + "," + Chr(34) + "KB4012216" + Chr(34) + "," + Chr(34) + "KB4012217" + Chr(34) + "," + Chr(34) + "KB4012598" + Chr(34) + "," + Chr(34) + "KB4012606" + Chr(34) + "," + Chr(34) + "KB4013198" + Chr(34) + "," + Chr(34) + "KB4013429" + Chr(34) + ", " + Chr(34) + "KB4022715" + Chr(34) + "" + vbCrLf)
        Script.Append("$hotfix = Get-HotFix -ComputerName $ComputerName | Where-Object {$hotfixes -contains $_.HotfixID} | Select-Object -expandproperty " + Chr(34) + "HotFixID" + Chr(34) + "" + vbCrLf)
        Script.Append("    if($hotfix) {" + vbCrLf)
        Script.Append("        $check = $True" + vbCrLf)
        Script.Append("                                $hf = $hotfix" + vbCrLf)
        Script.Append("    } " + vbCrLf)
        Script.Append("                else {" + vbCrLf)
        Script.Append("        $check = $False" + vbCrLf)
        Script.Append("                                $hf = " + Chr(34) + "None" + Chr(34) + "" + vbCrLf)
        Script.Append("    }" + vbCrLf)
        Script.Append("" + vbCrLf)
        Script.Append("   Write-Output (" + Chr(34) + "Computer Name : $thecomputername" + Chr(34) + ")" + vbCrLf)
        Script.Append("   Write-Output (" + Chr(34) + "IP Address    : $theipaddress" + Chr(34) + ")" + vbCrLf)
        Script.Append("   Write-Output (" + Chr(34) + "OS            : $theos" + Chr(34) + ")" + vbCrLf)
        Script.Append("   Write-Output (" + Chr(34) + "Computer Make : $thecomputermake" + Chr(34) + ")" + vbCrLf)
        Script.Append("   Write-Output (" + Chr(34) + "Computer Model: $thecomputermodel" + Chr(34) + ")" + vbCrLf)
        Script.Append("   Write-Output (" + Chr(34) + "CPU           : $thecpuname" + Chr(34) + ")" + vbCrLf)
        Script.Append("   Write-Output (" + Chr(34) + "Bios version  : $thebiosversion" + Chr(34) + ")" + vbCrLf)
        Script.Append("   Write-Output (" + Chr(34) + "CPU Load      : $thecpuloadaverage%" + Chr(34) + ")" + vbCrLf)
        Script.Append("   Write-Output (" + Chr(34) + "Memory Load   : $mem%" + Chr(34) + ")" + vbCrLf)
        Script.Append("   Write-Output (" + Chr(34) + "Last boot     : " + Chr(34) + " + $os.ConvertToDateTime($os.LastBootUpTime) )" + vbCrLf)
        Script.Append("   Write-Output (" + Chr(34) + "Uptime        : " + Chr(34) + " + $uptime.Days + " + Chr(34) + " Days " + Chr(34) + " + $uptime.Hours + " + Chr(34) + " Hours " + Chr(34) + " + $uptime.Minutes + " + Chr(34) + " Minutes" + Chr(34) + ")" + vbCrLf)
        Script.Append("   Write-Output (" + Chr(34) + "Wcry?         : " + Chr(34) + " + $check)" + vbCrLf)
        Script.Append("   Write-Output (" + Chr(34) + "Patch KB      : " + Chr(34) + " + $hf)" + vbCrLf)
        Script.Append("   " + vbCrLf)
        Script.Append("   " + vbCrLf)
        Script.Append("  " + vbCrLf)
        Script.Append(" " + vbCrLf)
        Script.Append("$theuptime = " + Chr(34) + "$($uptime.Days) Days $($uptime.Hours) Hours $($uptime.Minutes) Minutes $($uptime.Seconds) Seconds" + Chr(34) + "" + vbCrLf)
        Script.Append("$thememorypercent = " + Chr(34) + "$($mem) %" + Chr(34) + "" + vbCrLf)
        Script.Append("$user = Import-CSV " + Chr(34) + "C:\Program Files\Active Directory Admin Tool\Data\ticket_data.csv" + Chr(34) + "" + vbCrLf)
        Script.Append("$user.ComputerModel = $thecomputermodel" + vbCrLf)
        Script.Append("$user.ComputerSerial = $thecomputerserial" + vbCrLf)
        Script.Append("$user.OS = $theos" + vbCrLf)
        Script.Append("$user.CPU = $thecpuname" + vbCrLf)
        Script.Append("$user.MemoryTotal = $installedmemory" + vbCrLf)
        Script.Append("$user.LastBoot = $os.ConvertToDateTime($os.LastBootUpTime)" + vbCrLf)
        Script.Append("$user.WCryPatched = $check" + vbCrLf)
        Script.Append("$user.HotFix = $hf" + vbCrLf)
        Script.Append("$user.BiosVersion = $thebiosversion" + vbCrLf)
        Script.Append("$user.DiskFree = $disk0free" + vbCrLf)
        Script.Append("$user.DiskTotal = $disk0total" + vbCrLf)
        Script.Append("$user.IpAddress = $theipaddress" + vbCrLf)
        Script.Append("$user.SelectedComputer = $id" + vbCrLf)
        Script.Append("$user.ActiveUser = $theuser" + vbCrLf)
        Script.Append("$user.Uptime = $theuptime" + vbCrLf)
        Script.Append("$user.MemoryPercent = $thememorypercent" + vbCrLf)
        Script.Append("" + vbCrLf)
        Script.Append("$user | Export-CSV " + Chr(34) + "C:\Program Files\Active Directory Admin Tool\Data\ticket_data.csv" + Chr(34) + " -Force -NoTypeInformation" + vbCrLf)
        Script.Append("$items = " + Chr(34) + "C:\Program Files\Active Directory Admin Tool\Data\ticket_data.csv" + Chr(34) + "" + vbCrLf)
        Script.Append("(Get-Content $items) | Foreach-Object {$_ -replace '" + Chr(34) + "', " + Chr(34) + "" + Chr(34) + "} | Set-Content $items" + vbCrLf)
        Return Script.ToString()
    End Function

    'Get Members Script
    Private Function GetGroupMembers()
        Dim Script As New StringBuilder()
        Script.Append("Import-Module ActiveDirectory" + vbCrLf)
        Script.Append("Get-ADPrincipalGroupMembership " + TbUserID.Text.ToString + " | select name -Verbose" + vbCrLf)
        Return Script.ToString()
    End Function




    'Get apps script
    Private Function GetApps()
        Dim Script As New StringBuilder()
        Script.Append("# Get-InstalledApp.ps1" + vbCrLf)
        Script.Append("# Outputs installed applications" + vbCrLf)
        Script.Append("" + vbCrLf)
        Script.Append("param([String[]] $ComputerName," + vbCrLf)
        Script.Append("      [String] $AppID," + vbCrLf)
        Script.Append("      [String] $AppName," + vbCrLf)
        Script.Append("      [String] $Publisher," + vbCrLf)
        Script.Append("      [String] $Version," + vbCrLf)
        Script.Append("      [Switch] $MatchAll," + vbCrLf)
        Script.Append("      [Switch] $Help" + vbCrLf)
        Script.Append("     )" + vbCrLf)
        Script.Append("" + vbCrLf)
        Script.Append("$HKLM = [UInt32] " + Chr(34) + "0x80000002" + Chr(34) + "" + vbCrLf)
        Script.Append("$UNINSTALL_KEY = " + Chr(34) + "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" + Chr(34) + "" + vbCrLf)
        Script.Append("" + vbCrLf)
        Script.Append("# Outputs a usage message and ends the script." + vbCrLf)
        Script.Append("function usage {" + vbCrLf)
        Script.Append("  $scriptname = $SCRIPT:MYINVOCATION.MyCommand.Name" + vbCrLf)
        Script.Append("" + vbCrLf)
        Script.Append("  " + Chr(34) + "NAME" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "    $scriptname" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "SYNOPSIS" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "    Outputs installed applications on one or more computers that match one or" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "    more criteria." + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "SYNTAX" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "    $scriptname [-computername <String[]>] [-appID <String>]" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "    [-appname <String>] [-publisher <String>] [-version <String>] [-matchall]" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "PARAMETERS" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "    -computername <String[]>" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "        Outputs applications on the named computer(s). If you omit this" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "        parameter, the local computer is assumed." + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "    -appID <String>" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "        Select applications with the specified application ID. An application's" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "        appID is equivalent to its registry subkey in the location" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "        HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall. For Windows" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "        Installer-based applications, this is the application's product code" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "        GUID (e.g. {3248F0A8-6813-11D6-A77B-00B0D0160060})." + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "    -appname <String>" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "        Select applications with the specified application name. The appname is" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "        the application's name as it appears in the Add/Remove Programs list." + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "    -publisher <String>" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "        Select applications with the specified publisher name." + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "    -version <String>" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "        Select applications with the specified version." + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "    -matchall" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "        Output all matching applications instead of stopping after the first" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "        match." + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "NOTES" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "    All installed applications are output if you omit -appID, -appname," + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "    -publisher, and -version. Also, the -appID, -appname, -publisher, and" + Chr(34) + "" + vbCrLf)
        Script.Append("  " + Chr(34) + "    -version parameters all accept wildcards (e.g., -version 5.2.*)." + Chr(34) + "" + vbCrLf)
        Script.Append("" + vbCrLf)
        Script.Append("  exit" + vbCrLf)
        Script.Append("}" + vbCrLf)
        Script.Append("" + vbCrLf)
        Script.Append("function main {" + vbCrLf)
        Script.Append("  # If -help is present, output the usage message." + vbCrLf)
        Script.Append("  if ($Help) {" + vbCrLf)
        Script.Append("    usage" + vbCrLf)
        Script.Append("  }" + vbCrLf)
        Script.Append("" + vbCrLf)
        Script.Append("  # Create a hash table containing the requested application properties." + vbCrLf)
        Script.Append("  #CALLOUT A" + vbCrLf)
        Script.Append("  $propertyList = @{}" + vbCrLf)
        Script.Append("  if ($AppID -ne " + Chr(34) + "" + Chr(34) + ")     { $propertyList.AppID = $AppID }" + vbCrLf)
        Script.Append("  if ($AppName -ne " + Chr(34) + "" + Chr(34) + ")   { $propertyList.AppName = $AppName }" + vbCrLf)
        Script.Append("  if ($Publisher -ne " + Chr(34) + "" + Chr(34) + ") { $propertyList.Publisher = $Publisher }" + vbCrLf)
        Script.Append("  if ($Version -ne " + Chr(34) + "" + Chr(34) + ")   { $propertyList.Version = $Version }" + vbCrLf)
        Script.Append("  #END CALLOUT A" + vbCrLf)
        Script.Append("" + vbCrLf)
        Script.Append("  # Use the local computer's name if no computer name(s) specified." + vbCrLf)
        Script.Append("  if ($ComputerName -eq $NULL) {" + vbCrLf)
        Script.Append("    $ComputerName = " + Chr(34) + SelectedComputer.Text + Chr(34) + "" + vbCrLf)
        Script.Append("  }" + vbCrLf)
        Script.Append("" + vbCrLf)
        Script.Append("  # Iterate the computer name(s)." + vbCrLf)
        Script.Append("  foreach ($machine in $ComputerName) {" + vbCrLf)
        Script.Append("    $err = $NULL" + vbCrLf)
        Script.Append("" + vbCrLf)
        Script.Append("    # If WMI throws a RuntimeException exception," + vbCrLf)
        Script.Append("    # save the error and continue to the next statement." + vbCrLf)
        Script.Append("    #CALLOUT B" + vbCrLf)
        Script.Append("    trap [System.Management.Automation.RuntimeException] {" + vbCrLf)
        Script.Append("      set-variable err $ERROR[0] -scope 1" + vbCrLf)
        Script.Append("      continue" + vbCrLf)
        Script.Append("    }" + vbCrLf)
        Script.Append("    #END CALLOUT B" + vbCrLf)
        Script.Append("" + vbCrLf)
        Script.Append("    # Connect to the StdRegProv class on the computer." + vbCrLf)
        Script.Append("    #CALLOUT C" + vbCrLf)
        Script.Append("    $regProv = [WMIClass] " + Chr(34) + "\\$machine\root\default:StdRegProv" + Chr(34) + "" + vbCrLf)
        Script.Append("" + vbCrLf)
        Script.Append("    # In case of an exception, write the error" + vbCrLf)
        Script.Append("    # record and continue to the next computer." + vbCrLf)
        Script.Append("    if ($err) {" + vbCrLf)
        Script.Append("      write-error -errorrecord $err" + vbCrLf)
        Script.Append("      continue" + vbCrLf)
        Script.Append("    }" + vbCrLf)
        Script.Append("    #END CALLOUT C" + vbCrLf)
        Script.Append("" + vbCrLf)
        Script.Append("    # Enumerate the Uninstall subkey." + vbCrLf)
        Script.Append("    $subkeys = $regProv.EnumKey($HKLM, $UNINSTALL_KEY).sNames" + vbCrLf)
        Script.Append("    foreach ($subkey in $subkeys) {" + vbCrLf)
        Script.Append("      # Get the application's display name." + vbCrLf)
        Script.Append("      $name = $regProv.GetStringValue($HKLM," + vbCrLf)
        Script.Append("        (join-path $UNINSTALL_KEY $subkey), " + Chr(34) + "DisplayName" + Chr(34) + ").sValue" + vbCrLf)
        Script.Append("      # Only continue of the application's display name isn't empty." + vbCrLf)
        Script.Append("      if ($name -ne $NULL) {" + vbCrLf)
        Script.Append("        # Create an object representing the installed application." + vbCrLf)
        Script.Append("        $output = new-object PSObject                " + vbCrLf)
        Script.Append("        $output | add-member NoteProperty AppName -value $name" + vbCrLf)
        Script.Append("        # If the property list is empty, output the object;" + vbCrLf)
        Script.Append("        # otherwise, try to match all named properties." + vbCrLf)
        Script.Append("        if ($propertyList.Keys.Count -eq 0) {" + vbCrLf)
        Script.Append("          $output" + vbCrLf)
        Script.Append("        } else {" + vbCrLf)
        Script.Append("          #CALLOUT D" + vbCrLf)
        Script.Append("          $matches = 0" + vbCrLf)
        Script.Append("          foreach ($key in $propertyList.Keys) {" + vbCrLf)
        Script.Append("            if ($output.$key -like $propertyList.$key) {" + vbCrLf)
        Script.Append("              $matches += 1" + vbCrLf)
        Script.Append("            }" + vbCrLf)
        Script.Append("          }" + vbCrLf)
        Script.Append("          # If all properties matched, output the object." + vbCrLf)
        Script.Append("          if ($matches -eq $propertyList.Keys.Count) {" + vbCrLf)
        Script.Append("            $output" + vbCrLf)
        Script.Append("            # If -matchall is missing, break out of the foreach loop." + vbCrLf)
        Script.Append("            if (-not $MatchAll) {" + vbCrLf)
        Script.Append("              break" + vbCrLf)
        Script.Append("            }" + vbCrLf)
        Script.Append("          }" + vbCrLf)
        Script.Append("          #END CALLOUT D" + vbCrLf)
        Script.Append("        }" + vbCrLf)
        Script.Append("      }" + vbCrLf)
        Script.Append("    }" + vbCrLf)
        Script.Append("  }" + vbCrLf)
        Script.Append("}" + vbCrLf)
        Script.Append("" + vbCrLf)
        Script.Append("main" + vbCrLf)
        Script.Append("" + vbCrLf)
        Return Script.ToString()
    End Function

    'Garbage Collector, to collect the garbage you leave behind
    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    'Used to toggle wait cursor
    Public Sub Hourglass(ByVal Show As Boolean)
        If (Show = True) Then
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Else
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub


    'search user button!! this is where you search for the user in active directory
    Private Sub BtnSearch_Click(sender As System.Object, e As EventArgs) Handles BtnSearch.Click
        Hourglass(True) 'Calls the sub to turn on the wait cursor.
        ListView1.Items.Clear() 'Clear the items from the previous search.  

        Using Root As New DirectoryEntry 'Establish connection to current loged on users Active Directory
            Using Searcher As New DirectorySearcher(Root) 'Start at the top
                'If a first or last name isn't present search using the user ID field.
                If Not (TbFirst.Text.Length + TbLast.Text.Length) > 1 Then
                    Searcher.Filter = "(&(objectCategory=user)(ANR=" & TbUserID.Text & " * ))"
                Else
                    Searcher.Filter = "(&(objectCategory=user)(givenName=" & TbFirst.Text & "*" & ")(sn=" & TbLast.Text & "*" & "))"
                End If

                Searcher.SearchScope = SearchScope.Subtree 'Start at the top and keep drilling down
                Searcher.PropertiesToLoad.Add("sAMAccountName") 'Load User ID
                Searcher.PropertiesToLoad.Add("displayName") 'Load Display Name
                Searcher.PropertiesToLoad.Add("givenName") 'Load Users first name
                Searcher.PropertiesToLoad.Add("sn")   'Load Users last name
                Searcher.PropertiesToLoad.Add("distinguishedName")   'Users Distinguished name
                Searcher.Sort.PropertyName = "sAMAccountName" 'Sort by user ID
                Searcher.Sort.Direction = SortDirection.Ascending 'A-Z
                Using users = Searcher.FindAll 'Users contains our searh results
                    'MsgBox(users.Count)
                    If users.Count > 0 Then 'If it's zero then no matches were found
                        'Item 1 through Item 5 are the columns in our 1st listview control
                        Dim Item1 As String = Nothing 'User or Contact
                        Dim Item2 As String = Nothing 'sAMAccountName
                        Dim Item3 As String = Nothing 'givenName
                        Dim Item4 As String = Nothing 'sn
                        Dim Item5 As String = Nothing 'distinguishedName
                        Dim strDisplyName As String = Nothing

                        For Each User As SearchResult In users 'goes throug each user in the search results
                            If User.Properties.Contains("displayName") Then '<--This makes sure the property actually exists and has a value
                                strDisplyName = CStr(User.Properties("displayName").Item(0)) '<-- we need to use 0 here because this attribute only has one value

                            End If

                            Dim lv As ListViewItem = (ListView1.Items.Add(strDisplyName)) 'Display name goes in the first column of the listview.

                            If User.Properties.Contains("sAMAccountName") Then '<--This makes sure the property actually exists and has a value
                                Item1 = CStr(User.Properties("sAMAccountName").Item(0)) 'If it's got a sAMAcountName it's a user object.
                                Item1 = "User" 'The second column is user or contact
                            Else 'If there's no sAMAccountName I assume it's a contact
                                Item1 = "Contact" 'The second column user or contact
                            End If
                            lv.SubItems.Add(Item1) 'Add it to the listview
                            Try  'If this property is Null/Empty it will throw an exception. This traps it.
                                Item2 = CStr(User.Properties("sAMAccountName").Item(0))
                            Catch
                                Item2 = "" 'Since it's null set it to blank.
                            End Try
                            lv.SubItems.Add(Item2) 'Add it to the listview
                            Try  'If this property is Null/Empty it will throw an exception. This traps it.
                                Item3 = User.Properties("givenName").Item(0).ToString
                            Catch
                                Item3 = "" 'Since it's null set it to blank.
                            End Try
                            lv.SubItems.Add(Item3) 'Add it to the listview
                            Try  'If this property is Null/Empty it will throw an exception. This traps it.
                                Item4 = User.Properties("sn").Item(0).ToString
                            Catch
                                Item4 = "" 'Since it's null set it to blank.
                            End Try
                            lv.SubItems.Add(Item4) 'Add it to the listview
                            'Item 5
                            lv.SubItems.Add(User.Properties("distinguishedName").Item(0).ToString) 'Okay this item is added to the listview BUT thers's no column defind to display it. 
                            'I've found this to be very handy way of storing a value to use at a later point.
                        Next
                    Else

                        Exit Sub
                    End If
                End Using
            End Using
        End Using
        Hourglass(False)
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    'A line in the USER listview has been clicked!
    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged

        If ListView1.SelectedItems.Count = 1 Then 'I have to do this check or I'll get an index out of range on SelectedItems(0) in the following line.
            lblUserDN.Text = ListView1.SelectedItems(0).SubItems(5).Text 'I display the users DN in a label on the main form.
            'Note: I pulled this from the listviw data stored durring BtnSearch_Click sub above.
        End If

        Hourglass(True)
        ListView2.Items.Clear() 'Clear email addresses from previous selection
        'Establish connection to specific (The user or contact clicked on) AD object.
        Using user As New DirectoryEntry("LDAP://" & lblUserDN.Text) 'Here we get to use the data I pulled from the listviw durring BtnSearch_Click sub above.

            For Each eAddress As String In user.Properties("proxyAddresses") 'removes any occurances of matching eMail address.              
                Dim parts As String() = eAddress.Split(New Char() {":"c})
                AddItemToListView(parts(0), parts(1))
            Next

            If user.Properties.Contains("CanonicalName") Then '<--This makes sure the property actually exists and has a value
                Dim strUserCanonicalName = CStr(user.Properties("CanonicalName")(0)) '<-- we need to use 0 here because this attribute only has one value
            End If

            If user.Properties.Contains("description") Then '<--This makes sure the property actually exists and has a value
                Dim strDescription = CStr(user.Properties("description")(0)) '<-- we need to use 0 here because this attribute only has one value
                rtbDescription.Text = strDescription
            Else
                rtbDescription.Text = Nothing
            End If

            If user.Properties.Contains("sAMAccountName") Then '<--This makes sure the property actually exists and has a value
                TbUserID.Text = CStr(user.Properties("sAMAccountName")(0)) '<-- we need to use 0 here because this attribute only has one value


            Else
                TbUserID.Text = Nothing
            End If

            If user.Properties.Contains("department") Then '<--This makes sure the property actually exists and has a value
                lblCDepartment.Text = CStr(user.Properties("department")(0)) '<-- we need to use 0 here because this attribute only has one value
            End If

            If user.Properties.Contains("title") Then '<--This makes sure the property actually exists and has a value
                lblCRole.Text = CStr(user.Properties("title")(0)) '<-- we need to use 0 here because this attribute only has one value
            Else
                lblCRole.Text = Nothing
            End If

            'With these next few properties I get the values using a GetADProperty function I wrote. It's does the exact same thing as above.
            TbFirst.Text = GetADProperty(user, "givenName")
            TbLast.Text = GetADProperty(user, "sn")
            TbEmail.Text = GetADProperty(user, "mail")
            TbHome.Text = GetADProperty(user, "homePhone")
            TbMobile.Text = GetADProperty(user, "mobile")
            TbIPPhone.Text = GetADProperty(user, "ipPhone")
            TbAddress.Text = GetADProperty(user, "streetAddress")
            TbCity.Text = GetADProperty(user, "l")
            TbState.Text = GetADProperty(user, "st")
            TbZip.Text = GetADProperty(user, "postalCode")

            Dim strMgrDN As String = GetADProperty(user, "manager") 'The manager property is stored as the distinguished name to the manager. Here we go get it.
            If Not strMgrDN = "" Then  'Check if it's empty
                Using Manager As New DirectoryEntry("LDAP://" & strMgrDN) 'if it's not, set a directory entry point to it.
                    TbManager.Text = GetADProperty(Manager, "displayName") 'Get the display name of the manager and put it in the textbox.
                End Using 'close the connection to AD for this object.
            Else
                TbManager.Text = ""
            End If

            rtbNotes.Text = GetADProperty(user, "info")

            If user.Properties.Contains("thumbnailPhoto") Then '<--This makes sure the property actually exists and has a value
                Dim bytBLOBData() As Byte = CType((user.Properties("thumbnailPhoto")(0)), Byte()) '<-- we need to use 0 here because this attribute only has one value
                Using stmBLOBData As New MemoryStream(bytBLOBData) 'Create new memory stream.
                    pbUserImg.Image = Image.FromStream(stmBLOBData) 'Load image from stream
                End Using
            Else
                pbUserImg.Image = Nothing
                pbUserImg.Tag = "NoPicture"
            End If

            If ListView1.SelectedItems.Count = 1 Then 'I have to do this check or I'll get an index out of range on SelectedItems(0) in the following line.
                If ListView1.SelectedItems(0).SubItems(1).Text = "User" Then
                    lblConjunction.Text = "and"
                    'Check User Account control flags to check if user is disabled
                    If user.Properties.Contains("userAccountControl") Then
                        If IsAccountActive(CInt(user.Properties("userAccountControl")(0))) Then 'calls IsAccountActive function (See end )
                            lblAccountStatus.Text = " Enabled"
                            lblAccountStatus.ForeColor = System.Drawing.Color.Green 'Sets the text color
                        Else
                            lblAccountStatus.Text = "Disabled"
                            lblAccountStatus.ForeColor = System.Drawing.Color.Red 'Sets the text color
                        End If
                    End If
                    'Use System.DirectoryServices.AccountManagement to determine lockout status
                    Dim UserObject As AccountManagement.UserPrincipal = AccountManagement.UserPrincipal.FindByIdentity(New AccountManagement.PrincipalContext(AccountManagement.ContextType.Domain), Me.TbUserID.Text)
                    If UserObject.IsAccountLockedOut Then
                        lblLocked.Text = "Locked."
                        lblLocked.ForeColor = System.Drawing.Color.Red 'Sets the text color
                    Else
                        lblLocked.Text = "Unlocked."
                        lblLocked.ForeColor = System.Drawing.Color.Green 'Sets the text color
                    End If
                Else
                    lblAccountStatus.Text = "Contact"
                    lblAccountStatus.ForeColor = System.Drawing.Color.Black 'Sets the text color
                    lblLocked.Text = Nothing
                    lblConjunction.Text = Nothing
                End If
                'End use of System.DirectoryServices.AccountManagement to determin lockout status
            End If

            If user.Properties.Contains("homeMDB") Then 'if this is present than the account has a mailbox.
                lblHMB.Text = "YES"
            Else
                lblHMB.Text = "NO"
            End If
        End Using 'Close the directory entry to the user object

        Hourglass(False)
        'enable the AD management buttons
        BtnLockout.Enabled = True
        UnlockButton.Enabled = True

        Try
            TBADGroupMemberships.Text = RunScript(GetGroupMembers).ToString() 'runs powershell script to get user group membership
            TBADGroupMemberships.Text = TBADGroupMemberships.Text.Replace("@{name=", "")
            TBADGroupMemberships.Text = TBADGroupMemberships.Text.Replace("}", "")
        Catch
        End Try

        Hourglass(True)
        GC.Collect()
        GC.WaitForPendingFinalizers()
        Hourglass(False)
        'cleanup


        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    'Add Exchange info to List
    Private Sub AddItemToListView(ByVal Item As String, ByVal Item2 As String)
        Dim lv As ListViewItem = (ListView2.Items.Add(Item))
        lv.SubItems.Add(Item2)
    End Sub

    'Checks if the Account is Active
    Public Shared Function IsAccountActive(ByVal userAccountControl As Integer) As Boolean
        Dim flagExists As Integer = userAccountControl And &H2 'This does a binary AND of userAccountControl and 2 if the flag is set the outcome is 0 if not it's 2
        'if a match is found, then the disabled flag exists within the control flags
        If flagExists > 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    'Gets Ad Properties of User
    Public Shared Function GetADProperty(ByVal de As DirectoryEntry, ByVal pName As String) As String
        Dim pValue As String
        Try
            pValue = de.Properties(pName).Value.ToString() 'When value is found return it. . .
        Catch
            pValue = "" 'When property dosn't exist set value to null and return . . .
            'MsgBox("Property Notfound =" & pName)
        End Try
        Return (pValue) '. . .here
    End Function

    'Press Enter in Computername Field to search
    Private Sub SelectedComputer_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles SelectedComputer.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            Call SearchComputer_Click(sender, e)
        End If
    End Sub

    'Press Enter in UserID Field to search
    Private Sub TbUserID_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles TbUserID.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            TbFirst.Clear()
            TbLast.Clear()
            Call BtnSearch_Click(sender, e)
        End If
    End Sub

    'Press Enter in Firstname
    Private Sub TbFirst_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles TbFirst.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            Call BtnSearch_Click(sender, e)
        End If
    End Sub

    'Press Enter in Lastname
    Private Sub TbLast_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles TbLast.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            Call BtnSearch_Click(sender, e)
        End If
    End Sub

    'New User Button
    Private Sub NewUser_Click(sender As Object, e As EventArgs) Handles NewUser.Click
        TbUserID.Text = ""
        TbFirst.Text = ""
        TbLast.Text = ""
        TbEmail.Clear()
        TbHome.Clear()
        TbMobile.Clear()
        TbIPPhone.Clear()
        TbAddress.Clear()
        TbCity.Clear()
        TbState.Clear()
        TbZip.Clear()
        SelectedComputer.Clear()
        rtbDescription.Clear()
        rtbNotes.Clear()
        ListView1.Items.Clear()
        ListView2.Items.Clear()
        TbManager.Clear()
        TbUserID.Focus()
        ModelValue.Clear()
        SerialValue.Clear()
        OSValue.Clear()
        CPUValue.Clear()
        MemoryValue.Clear()
        HDFreeValue.Clear()
        HDTotalValue.Clear()
        SubnetValue.Clear()
        LastBootValue.Clear()
        WcryValue.Clear()
        LastBootValue.Clear()
        TbUptime.Clear()
        WcryValue.Clear()
        KBValue.Clear()
        TbMemoryPercent.Clear()
        TbManager.Clear()
        TbActiveUser.Clear()
        lblHMB.Text = ""
        lblAccountStatus.Text = ""
        lblConjunction.Text = ""
        lblLocked.Text = ""
        lblCRole.Text = ""
        lblCDepartment.Text = ""
        ResultsBox.Clear()
        TBADGroupMemberships.Clear()
        BtnLockout.Enabled = False
        UnlockButton.Enabled = False


    End Sub

    'Unlock Button
    Private Sub UnlockButton_Click(sender As Object, e As EventArgs) Handles UnlockButton.Click
        Hourglass(True)
        Dim RootDSE As New DirectoryEntry("LDAP://RootDSE")
        Dim DomainDN As String = RootDSE.Properties("DefaultNamingContext").Value
        Dim ADEntry As New DirectoryEntry("LDAP://" & DomainDN)
        Dim ADSearch As New DirectorySearcher(ADEntry)
        Dim UserID As String = TbUserID.Text

        ADSearch.Filter = ("(samAccountName=" & UserID & ")")
        ADSearch.SearchScope = SearchScope.Subtree
        Dim UserFound As SearchResult = ADSearch.FindOne()
        If Not IsNothing(UserFound) Then

            Dim Attrib As String = "msDS-User-Account-Control-Computed"
            Dim User As DirectoryEntry
            User = UserFound.GetDirectoryEntry()
            User.RefreshCache(New String() {Attrib})
            Const UF_LOCKOUT As Integer = &H10
            Dim Flags As Integer = CInt(Fix(User.Properties(Attrib).Value))

            If Convert.ToBoolean(Flags And UF_LOCKOUT) Then

                'Unlock account
                User.Properties("LockOutTime").Value = 0
                User.CommitChanges()
                lblLocked.Text = "Unlocked."
                lblLocked.ForeColor = Color.Green
            Else
            End If
        End If
        Hourglass(False)
    End Sub

    'Exit Application from File menu
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Close()
    End Sub

    'Copy computer details to clipboard button
    Private Sub BtnCopyDetails_Click(sender As Object, e As EventArgs) Handles BtnCopyDetails.Click
        Dim str As String
        str = "Active User: " & TbActiveUser.Text & vbCrLf &
            "Computer: " & SelectedComputer.Text & vbCrLf &
            "Model: " & ModelValue.Text & vbCrLf &
            "Serial: " & SerialValue.Text & vbCrLf &
            "OS: " & OSValue.Text & vbCrLf &
            "CPU: " & CPUValue.Text & vbCrLf &
            "Memory: " & MemoryValue.Text & vbCrLf &
            "HD Free: " & HDFreeValue.Text & vbCrLf &
            "HD Total: " & HDTotalValue.Text & vbCrLf &
            "IP: " & SubnetValue.Text & vbCrLf &
            "Last boot: " & LastBootValue.Text & vbCrLf &
            "Uptime: " & TbUptime.Text & vbCrLf &
            "Patched: " & WcryValue.Text & " " & KBValue.Text
        Clipboard.SetText(str)
    End Sub

    'Copy info to clipboard button
    Private Sub BtnCopyInfo_Click(sender As Object, e As EventArgs) Handles BtnCopyInfo.Click
        Dim str As String
        str = "Name: " & TbFirst.Text & " " & TbLast.Text & vbCrLf &
            "Email: " & TbEmail.Text & vbCrLf &
            "Home Phone: " & TbHome.Text & vbCrLf &
            "Mobile: " & TbMobile.Text & vbCrLf &
            "IP Phone: " & TbIPPhone.Text & vbCrLf &
            "Address: " & TbAddress.Text & " " & TbCity.Text & " " & TbState.Text & " " & TbZip.Text
        Clipboard.SetText(str)
    End Sub
    Private Sub BtnLockout_Click(sender As Object, e As EventArgs) Handles BtnLockout.Click
        ResultsBox.Text = RunScript(Lockoutstatus).ToString()
    End Sub

    'LockoutStatus Powershell Script
    Private Function Lockoutstatus()
        Dim Script As New StringBuilder()
        Dim EmployeeName As String = TbUserID.Text
        Script.Append("$EmployeeNumber = " + Chr(34) + EmployeeName + Chr(34) + vbCrLf)
        Script.Append("$program = " + Chr(34) + "C:\Program Files\Active Directory Admin Tool\Apps\lockoutstatus.exe" + Chr(34) + "" + vbCrLf)
        Script.Append("$programArgs = " + Chr(34) + "-u:$EmployeeNumber@$env:USERDNSDOMAIN" + Chr(34) + "" + vbCrLf)
        Script.Append("Invoke-Command -ScriptBlock { & $program $programArgs }" + vbCrLf)
        Return Script.ToString()
    End Function



    'Search Computer button - gets computer info using powershell
    Private Sub SearchComputer_Click(sender As Object, e As EventArgs) Handles SearchComputer.Click
        Hourglass(True)
        Try
            ResultsBox.Text = RunScript(GetUserComputerSelected).ToString() 'runs powershell script to get computer info

            Using sr As New StreamReader("C:\Program Files\Active Directory Admin Tool\Data\ticket_data.csv")
                sr.ReadLine.Split(","c)
                While Not sr.EndOfStream
                    Dim newline() As String = sr.ReadLine.Split(","c)
                    ModelValue.Text = (newline(27)) 'Model of Computer
                    SerialValue.Text = (newline(26)) 'Serial of Computer
                    OSValue.Text = (newline(13)) 'OS value
                    CPUValue.Text = (newline(14)) 'CPU Value
                    MemoryValue.Text = (newline(15)) + " GB" 'Memory in GB
                    LastBootValue.Text = (newline(16)) 'Last Boot up time
                    WcryValue.Text = (newline(19)) 'wcry patched true or false
                    KBValue.Text = (newline(18)) 'wcry kb#
                    HDFreeValue.Text = (newline(20)) + " GB" 'Total free space in GB
                    HDTotalValue.Text = (newline(21)) + " GB" 'Total space in GB
                    SubnetValue.Text = (newline(22)) 'IP Address
                    TbActiveUser.Text = (newline(25)) 'Active User
                    TbUptime.Text = (newline(24)) 'Uptime
                    TbMemoryPercent.Text = (newline(23)) + " in use" 'Memory%
                End While
                sr.Close()
            End Using
        Catch
        End Try
        GC.Collect()
        GC.WaitForPendingFinalizers()
        Hourglass(False)
    End Sub

    'Donate button
    Private Sub DonateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DonateToolStripMenuItem.Click
        Process.Start("https://www.paypal.me/rm519")
    End Sub

    'LinkedIn button
    Private Sub CreatedByArmanRamazyanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreatedByArmanRamazyanToolStripMenuItem.Click
        Process.Start("http://www.linkedin.com/in/arman-ramazyan")
    End Sub

    'Email me button
    Private Sub EmailArmanRamazyangmailcomToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EmailArmanRamazyangmailcomToolStripMenuItem.Click
        Dim Outl As Object
        Outl = CreateObject("Outlook.Application")
        If Outl IsNot Nothing Then
            Dim omsg As Object
            omsg = Outl.CreateItem(0)
            omsg.To = "Arman.Ramazyan@gmail.com"
            omsg.subject = ""
            omsg.body = ""
            'set message properties here...'
            omsg.Display(True) 'will display message to user
        End If
    End Sub

    'Twitter button
    Private Sub TwitterRm519ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TwitterRm519ToolStripMenuItem.Click
        Process.Start("https://twitter.com/rm519")
    End Sub


    Private Sub DonateToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        Process.Start("https://www.paypal.me/rm519/")
    End Sub


    Private Sub GetAppsButton_Click(sender As Object, e As EventArgs) Handles GetAppsButton.Click
        Hourglass(True)
        InstalledApps.Show()

        Try
            InstalledApps.InstalledAppsRtb.Text = RunScript(GetApps).ToString() 'runs powershell script to get apps
            InstalledApps.InstalledAppsRtb.Text = InstalledApps.InstalledAppsRtb.Text.Replace("@{AppName=", "")
            InstalledApps.InstalledAppsRtb.Text = InstalledApps.InstalledAppsRtb.Text.Replace("}", "")
        Catch
        End Try
        Hourglass(False)
    End Sub

    Private Sub WebsiteToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles WebsiteToolStripMenuItem.Click
        Process.Start("http://www.rm519.com")
    End Sub

    Private Sub ListGroupMembersButton_Click(sender As Object, e As EventArgs)
        Hourglass(True)

        Try
            TBADGroupMemberships.Text = RunScript(GetGroupMembers).ToString() 'runs powershell script to get user group memberships
        Catch
        End Try
        Hourglass(False)
    End Sub

End Class