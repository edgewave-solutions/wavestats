' WaveStats.vbs
' Affiche des statistiques système basiques : CPU, RAM et espace disque

Option Explicit

Dim objWMIService, colItems, objItem
Dim strComputer, strMsg
strComputer = "."

' Connexion à WMI
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")

' Informations sur le processeur
Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")
For Each objItem in colItems
    strMsg = "Processeur : " & objItem.Name & vbCrLf
Next

' Informations sur la mémoire RAM
Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
For Each objItem in colItems
    strMsg = strMsg & "Mémoire totale : " & FormatNumber(objItem.TotalVisibleMemorySize/1024,2) & " MB" & vbCrLf
    strMsg = strMsg & "Mémoire libre : " & FormatNumber(objItem.FreePhysicalMemory/1024,2) & " MB" & vbCrLf
Next

' Informations sur le disque
Set colItems = objWMIService.ExecQuery("Select * from Win32_LogicalDisk Where DriveType=3")
For Each objItem in colItems
    strMsg = strMsg & "Lecteur " & objItem.DeviceID & " : " & FormatNumber(objItem.FreeSpace/1073741824,2) & " GB libre sur " & FormatNumber(objItem.Size/1073741824,2) & " GB" & vbCrLf
Next

' Affiche les statistiques
MsgBox strMsg, vbInformation, "WaveStats - Statistiques système"
