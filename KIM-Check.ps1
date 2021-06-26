# --- Parameter zur Überprüfung festlegen ---
powershell -command "[console]::WindowWidth=120; [console]::WindowHeight=50; [console]::BufferWidth=[console]::WindowWidth"

$ports_to_check = 8995, 8465, 9999, 4443, 995, 465

$services_to_check = , "CGM_KIM_ClientModule", "kv.dox KIM Clientmodul Service"

$ms_path = "D:\medistar\"
$sys_path = "C:\Windows\SysWOW64\sysconf.s"

# --- Ab hier müssen keine Änderungen mehr vergenommen werden ---


$greenCheck = @{
    Object          = [Char]8730
    ForegroundColor = 'Green'
    NoNewLine       = $true
}

$redCross = @{
    Object          = 'X'
    ForegroundColor = 'Red'
    NoNewLine       = $true
}

function PrintBanner {
    Write-Host "  _  _____ __  __        ____ _               _    "
    Write-Host " | |/ /_ _|  \/  |      / ___| |__   ___  ___| | __"
    Write-Host " | ' / | || |\/| |_____| |   | '_ \ / _ \/ __| |/ /"
    Write-Host " | . \ | || |  | |_____| |___| | | |  __/ (__|   < "
    Write-Host " |_|\_\___|_|  |_|      \____|_| |_|\___|\___|_|\_\"
}
function Show-Icon {
    param (
        $icon
    )

    if ($icon -eq "success") {
        Write-Host "  [" -NoNewline
        Write-Host @greenCheck
        Write-Host "] " -NoNewline
    }
    else {
        Write-Host "  [" -NoNewline
        Write-Host @redCross
        Write-Host "] " -NoNewline
    }
    
}

function CheckUNC {
    param(
        $servicepath
    )

    $currentDirectory = Resolve-Path $servicepath
    $currentDrive = Split-Path -qualifier $currentDirectory.Path
    $logicalDisk = Get-WmiObject Win32_LogicalDisk -filter "DriveType = 4 AND DeviceID = '$currentDrive'"
    $uncPath = $currentDirectory.Path.Replace($currentDrive, $logicalDisk.ProviderName)
    if ($uncPath.Substring(0, 2) -eq "\\") {
        return $true
    }
    else {
        return $false
    }
}


function Portcheck {
    param(
        $port
    )

    #Port überprüfen
    $r = Get-NetTCPConnection | Where-Object Localport -eq $port | Select-Object -ExpandProperty OwningProcess
    if ($null -eq $r) {

        #Port wird von keinem Prozess genutzt => Erfolg
        Show-Icon "success"
        Write-Host "Port $port ist frei"
    }
    else {
        $owner_pid = $r[0]
        $owner_processname = Get-Process -Id $owner_pid | Select-Object -ExpandProperty ProcessName
        
        if (($owner_processname -Match "javaw") -or ($owner_processname -Match "KIM.ClientModul.ApplicationService")) {

            #Port wird bereits vom ClientModule-Prozess genutzt => Erfolg
            Show-Icon "success"
            Write-Host "Port $port wird verwendet von $owner_processname (PID: $owner_pid)"
        }
        else {

            #Port wird von Drittprozess blockiert => Problem
            Show-Icon "error"
            Write-Warning "Port $port wird verwendet von $owner_processname (PID: $owner_pid)"
        }
        
    }

}

function Servicecheck {
    param(
        $servicename
    )

    # Services abrufen, die ClientModule im Namen haben:
    $r = Get-Service "*$servicename*"
    if ($r.length -eq 0) {

        # Kein ClientModul-Dienst gefunden => Problem
        Show-Icon "error"
        Write-Warning "Dienst $servicename existiert nicht"
        return
    }

    foreach ($service in $r) {
        # Wenn ClientModul-Dienste gefunden werden: darüber iterieren
        $checkedServiceStatus = $service | Select-Object -ExpandProperty Status
        $checkedServiceName = $service | Select-Object -ExpandProperty DisplayName

        if ($checkedServiceStatus -eq "Running") {
            # Dienst ist aktiv => Erfolg
            Show-Icon "success"
            Write-Host "Dienst $checkedServiceName ist aktiv"

            #Auf Netzwerkpfad überprüfen:
            $servicePath = Get-CimInstance -ClassName win32_service | Where-Object { $_.Name -match '^CGM_KIM' } | Select-Object -ExpandProperty PathName
            if (CheckUNC $servicePath) {
                Show-Icon "error"
                Write-Warning "Dienst-Pfad ist ein UNC-Pfad: $servicePath"
            }
            else {
                Show-Icon "success"
                Write-Host "Dienst-Pfad ist kein UNC-Pfad"

            }
        }
        else {
            # Dienst ist nicht aktiv => Problem
            Show-Icon "error"
            Write-Warning "Dienst $checkedServiceName hat den Status: $checkedServiceStatus"
        }
    }
}



function Plugin {
    $plugin_config = "para\msinclude\globalvariable\CGMCONNECT_CONFIGS\KOMLEPlugin\config.xml"
    $fpath = Join-Path -path $ms_path -ChildPath $plugin_config

    #Prüfen, ob Datei existiert
    if (! (Test-Path -path $fpath)) {
        Show-Icon "error"
        Write-Warning "Keine KOM-LE-Konfiguration gefunden: $fpath"
        return
    }

    Show-Icon "success"
    Write-Host "KOM-LE-Konfiguration vorhanden"

    # XML-Datei einlesen
    [XML]$connect = Get-Content $fpath

    #Client-Adresse auslesen und prüfen, ob sie identisch zum Computernamen ist
    $clientAdresse = $connect.GeneralConfiguration.komLeClientAdresse
    if ($clientAdresse) {

        if ($env:computername -eq $clientAdresse) {
            Show-Icon "success"
            Write-Host " - Die 'KOM-LE ClientAdresse' entspricht dem Computernamen: '$clientAdresse'"
        }
        else {
            Show-Icon "error"
            Write-Warning " - Die 'KOM-LE ClientAdresse' ('$clientAdresse') entspricht nicht dem Computernamen ('$env:computername')"
        }
        
    }
    else {
        Show-Icon "error"
        Write-Warning " - Keine 'Client-Adresse' angegeben"
    }

    # Fachdienstadresse auslesen und prüfen
    $fachdienstAdresse = $connect.GeneralConfiguration.Fachdienstadresse
    if ($fachdienstAdresse) {
        Show-Icon "success"
        Write-Host " - Fachdienstadresse: '$fachdienstAdresse'"
        
    }
    else {
        Show-Icon "error"
        Write-Warning " - Keine 'Fachdienstadresse' angegeben"
    }

    # LDAP-URL auslesen und prüfen
    $ldapUrl = $connect.GeneralConfiguration.ldapUrl
    if ($ldapUrl) {
        Show-Icon "success"
        Write-Host " - LDAP-URL: '$ldapUrl'"
        
    }
    else {
        Show-Icon "error"
        Write-Warning " - Keine 'ldapUrl' angegeben"
    }

    # Ports auslesen und prüfen
    $pop3Port = $connect.GeneralConfiguration.pop3Port
    $smtpPort = $connect.GeneralConfiguration.smtpPort
    $komLeClientManagementPort = $connect.GeneralConfiguration.komLeClientManagementPort
    if ($pop3Port -and $smtpPort -and $komLeClientManagementPort) {
        Show-Icon "success"
        Write-Host " - Ports (POP3, SMTP, Management): $pop3Port, $smtpPort, $komLeClientManagementPort"
    }
    else {
        Show-Icon "error"
        Write-Warning " - Es sind nicht alle Ports (POP3, SMTP, Management) angegeben"
    }
}

function Secret {
    $sec = "KIM\KIM_Clientmodul\conf\*.sec"
    $spath = Join-Path -path $ms_path -ChildPath $sec

    #Prüfen, ob Datei existiert
    if (! (Test-Path -path $spath)) {
        Show-Icon "error"
        Write-Warning "Keine Secret-Datei gefunden: $spath"
        return
    }

    Show-Icon "success"
    Write-Host "Secret-Datei vorhanden"
}

function sysconf {
    #Test ob Datei vorhanden
    if(!(Test-Path -path $sys_path)){
    Show-Icon "error"
    Write-Warning "Keine sysconf.s gefunden: $sys_path"
    return
    }

    Show-Icon "success"
    Write-Host "sysconf.s vorhanden"

    Copy-Item -Path $sys_path -Destination "d:\medistar\sysconf.txt"
    
    #sysconf auslesen
    $content = Get-Content "d:\medistar\sysconf.txt" | Where-Object {$_ -like "*MS4 = d:\MEDISTAR\para*"}

    #Prüfen ob der UNC-Pfad hinterlegt ist
    #$ms4 = Get-Content $content

    if ($content -like "*MS4 = d:\MEDISTAR\para*"){
        Show-Icon "success"
        Write-Host "lokaler Pfad in der sysconf.s"
    }
    else {
        Show-Icon "error"
        Write-Warning "UNC-Pfad, bitte korrigieren in d:\medistar"
    }
}

function admin{
        
    $role = whoami /groups /fo csv | convertfrom-csv | where-object { $_.SID -eq "S-1-5-32-544" }

    if ($role -like "*Administratoren*"){
        Show-Icon "success"
        Write-Host "Nutzer hat administrative Rechte"
    }else{
        Show-Icon "error"
        Write-Warning "Nutzer hat keine administrativen Rechte"
    }
}

function dbms{

    $exe = "\prg4\m42t.exe"    
    $version = (Get-Item (Join-Path -path $ms_path -Childpath $exe)).VersionInfo.ProductVersion

    if ($version -ge "404.76"){
        Show-Icon "success"
        Write-Host "Medistar ist aktuell: $version"
    }else{
        Show-Icon "error"
        Write-Warning "Medistar ist nicht aktuell: $version"
    }
}

#Banner anzeigen:
PrintBanner
Write-Host ""
Write-Host ""

$inp = Read-Host -Prompt "[v]or der Installation oder [d]anach?"

if ($inp -eq "v"){

    #Windowsnutzer testen
    Write-Host ""
    Write-Host "  administrative Rechte"
    Write-Host "  ---------------------------------"
    admin
    Write-Host ""

    #Medistarversion testen
    Write-Host ""
    Write-Host "  Medistarversion testen"
    Write-Host "  ---------------------------------"
    dbms
    Write-Host ""

    #Ports checken:
    Write-Host ""
    Write-Host "  Checken der Ports"
    Write-Host "  ---------------------------------"
    foreach ($port in $ports_to_check) {
    Portcheck $port
    }
    Write-Host ""

    #sysconf.s checken
    Write-Host ""
    Write-Host "  Checken der sysconf"
    Write-Host "  ---------------------------------"
    sysconf
    Write-Host ""
    pause
}else{
    #Services checken:
    Write-Host ""
    Write-Host "  Status Windowsdienst"
    Write-Host "  ---------------------------------"
    foreach ($service in $services_to_check) {
       Servicecheck $service
    }
    Write-Host ""

    #Ports checken:
    Write-Host ""
    Write-Host "  Checken der Ports"
    Write-Host "  ---------------------------------"
    foreach ($port in $ports_to_check) {
    Portcheck $port
    }
    Write-Host ""

    #Plugin checken:
    Write-Host ""
    Write-Host "  Connect KIM-Plugin Konfiguration"
    Write-Host "  ---------------------------------"
    Plugin
    Write-Host ""

    #Secret checken:
    Write-Host ""
    Write-Host "  Secret Datei"
    Write-Host "  ---------------------------------"
    Secret
    Write-Host ""
    pause
}