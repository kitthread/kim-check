# --- Parameter zur Überprüfung festlegen ---
powershell -command Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser
powershell -command "[console]::WindowWidth=120; [console]::WindowHeight=50; [console]::BufferWidth=[console]::WindowWidth"

$ports_to_check_cgm = 8995, 8465, 9999, 4443 
$ports_to_check_kbv = 4443, 995, 465

$services_to_check_cgm = , "CGM_KIM_ClientModule"
$services_to_check_kbv = , "kv.dox KIM Clientmodul Service"

$ms_path = $ENV:medistardir
$sys_path = "C:\Windows\SysWOW64\sysconf.s"
$appdata = $ENV:APPDATA

$ms_version = "404.86"
$connect_version = "2.2.0.1"

$certificates_to_check = "KIM\KIM_Assist\data\kim\data\*.p12"
$certificates_to_check2 = "KIM\KIM_Assist\data\kim\data\"

$reg_key_java_path = "Registry::HKLM\SOFTWARE\WOW6432Node\CompuGROUP\Java Runtime Environment"
$javawX86 = "C:\Program Files (x86)\CG\Java\bin\javaw.exe"
$javawC = "C:\CG\Java\bin\javaw.exe"

$kim_assist_path = "KIM\KIM_Assist\"
$kim_assist_current_version = "KIM-Einrichtung-Assistent.jar"
$kim_assist_old_version = "KIM-Einrichtung-Assistent-1.0.14.jar"

$kim_client_path = "KIM\KIM_Clientmodul\"
$kim_client_current_version = "KIM-CM-10.0.2-15.jar"
$kim_client_old_version = "KIM-CM-10.0.2-14.jar"

$scriptversion = "2.0" # to use func updateS uncomment line 750

# --- Ab hier müssen keine Änderungen mehr vergenommen werden ---


$greenCheck = @{
    Object          = [Char]8730
    ForegroundColor = 'DarkGreen'
    NoNewLine       = $true
}

$redCross = @{
    Object          = 'X'
    ForegroundColor = 'Red'
    NoNewLine       = $true
}

function Set-ConsoleColor ($bc, $fc) {
    $Host.UI.RawUI.BackgroundColor = $bc
    $Host.UI.RawUI.ForegroundColor = $fc
    Clear-Host
}
Set-ConsoleColor 'gray' 'black'

function PrintBanner {
    Write-Host "                        "
    Write-Host "                   ))) "
    Write-Host "                  ((( "
    Write-Host "                +-----+ "
    Write-Host "                !     !] "
    Write-Host "                 `-----' " 
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

function updateS{

	$http_request = [System.Net.WebRequest]::Create('https://www.memski.org/2.0.html')
    try {
        $http_response = $http_request.GetResponse()

	    $http_status = [int]$http_response.StatusCode

	    If ($http_status -eq 200) {
   		    Write-Host ""
	    }else{
    	    Write-Host "The Site may be down, please check!"
	    }        
    }
    catch {
        Write-Host ""
        Write-Warning "Das Script ist nicht mehr aktuell."
        Write-Warning "Bitte die aktuelle Version nutzen."
        
        $source = "https://www.memski.org/kim_check_2.0s.zip"
        $dest = $ms_path+"\archiv\" + $(Split-Path -Path $source -Leaf)
        
        Invoke-WebRequest -Uri $source -OutFile $dest -UseBasicParsing
        Write-Host ""
        Write-Host ""
        Write-Host ""
        Show-Icon "success"
        Write-Host "Die Aktuelle Version wurde heruntergeladen, liegt im Ordner Archiv :)"

        Pause
        exit
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

function KimAssistInstallationCheck {
	Write-Host "  KIM Assist"
    # Prüfen, ob KIM Assist existiert
	$assist_path = Join-Path -path $ms_path -ChildPath $kim_assist_path
    if (! (Test-Path -path $assist_path)) {
        Show-Icon "error"
        Write-Host "Keine KIM Assist Installation gefunden: $assist_path"
        return
    }
	
	$kim_assist_lib = "data\kim\lib"
	$libpath = Join-Path -path $assist_path -ChildPath $kim_assist_lib
	
	# Prüfen, ob aktuelle KIM Assist Version existiert
	$assist_cur_version = Join-Path -path $libpath -ChildPath $kim_assist_current_version
	if (! (Test-Path -path $assist_cur_version) ) {
		Show-Icon "error"
        Write-Host "Keine aktuelle KIM Assist Version gefunden: $assist_cur_version"
	} else {
		Show-Icon "success"
        Write-Host "Aktuelle KIM Assist Version gefunden: $assist_cur_version"
	}
	
	#Prüfen, ob alte KIM Assist Version existiert
    $kim_old_version = Join-Path -path $libpath -ChildPath $kim_assist_old_version
	if ( (Test-Path -path $kim_old_version) ) {
		Show-Icon "error"
        Write-Host "Alte KIM Assist Version gefunden, bitte loeschen: $kim_old_version"
	}
}

function KimClientmodulInstallationCheck {
	Write-Host ""
	Write-Host "  KIM Clientmodul"
	
	# Prüfen, ob KIM Clientmodul existiert
    $client_path = Join-Path -path $ms_path -ChildPath $kim_client_path
    if (! (Test-Path -path $client_path)) {
        Show-Icon "error"
        Write-Host "Keine KIM Assist Installation gefunden: $client_path"
        return
    }
	
    $kim_client_lib = "libs\common"
	$commonpath = Join-Path -path $client_path -ChildPath $kim_client_lib
	
	# Prüfen, ob aktuelle KIM Clientmodul Version existiert
	$client_cur_version = Join-Path -path $commonpath -ChildPath $kim_client_current_version
	if (! (Test-Path -path $client_cur_version) ) {
		Show-Icon "error"
        Write-Warning "Keine aktuelle KIM Clientmodul Version gefunden:"
        Show-Icon "error"
        Write-Warning "$client_cur_version"

	} else {
		Show-Icon "success"
        Write-Host "Aktuelle KIM Clientmodul Version gefunden: $client_cur_version"
	}
	
	#Prüfen, ob alte KIM Clientmodul Version existiert
    $client_old_version = Join-Path -path $commonpath -ChildPath $kim_client_old_version
	if ( (Test-Path -path $client_old_version) ) {
		Show-Icon "error"
        Write-Warning "Alte KIM Clientmodul Version gefunden, bitte loeschen:"
        Show-Icon "error"
        Write-Warning "$client_old_version"
	}
}


function MedistarJavaPathCheck {

    # Prüfen ob Registry-Key existiert
    if ((Test-Path $reg_key_java_path) -eq $false) {
        Show-Icon "error"
        Write-Host "MEDISTAR-InstallationsPfad nicht in Registry ($reg_key_java_path)"
        return $false
    }

    # Registry-Key auslesen
    $MEDISTAR_JAVA_PFAD = Get-ItemProperty -Path "Registry::HKLM\SOFTWARE\WOW6432Node\CompuGROUP\Java Runtime Environment" -Name CurrentPath -ErrorAction SilentlyContinue
    if (! $MEDISTAR_JAVA_PFAD) {
        Show-Icon "error"
        Write-Host "MEDISTAR-Javapfad ist nicht in Registry ($MEDISTAR_JAVA_PATH)"
        return $false
    }

	$Javapfad = $MEDISTAR_JAVA_PFAD.CurrentPath	
	
	# Prüfen ob Java-Verzeichnis existiert
	if ( !(Test-Path -path $Javapfad) ) {
		Show-Icon "error"
		Write-Host "Medistar-Javaverzeichnis ist nicht vorhanden: $Javapfad"
		Write-Host "      Bzw. der Registry-Schluessel muss angepasst werden: HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CompuGROUP\Java Runtime Environment\CurrentPath"
		return
	}

	if ( $Javapfad -Match "\(x86\)" ) {
		if( Test-Path $javawX86){
            Show-Icon "success"
            Write-Host "Java-Installationspfad und Registry-Eintrag stimmen ueberein (x86)."
        } else{
            Show-Icon "error"
            Write-Warning "Java-Installationspfad und Registry-Eintrag stimmen nicht ueberein."
        }
	}else {
        if( Test-Path $javawC){
            Show-Icon "success"
            Write-Host "Java-Installationspfad und Registry-Eintrag stimmen ueberein (C:\)"
        }else{
            Show-Icon "error"
            Write-Warning "Java-Installationspfad und Registry-Eintrag stimmen nicht ueberein."
        }
	}

}

function RESTCheck {
    
    $confprop = "KIM\KIM_Clientmodul\conf\cm-config.properties"
    $confPath = Join-Path -path $ms_path -ChildPath $confprop
    $portmgmt = ""

    if (! (Test-Path -path $confPath)){
        Show-Icon "error"
        Write-Warning "Keine cm-config.properties gefunden: $confPath"
    }

    if ((Get-Content $confPath) -Match "PORT_MGMT=4443"){
        $portmgmt = 4443
    }else{
        $portmgmt = 9999
    }

    try {
    #Prüfen, ob REST-Schnitstelle als Status "OK" zurückgibt
    $r = Invoke-WebRequest -URI "http://localhost:$portmgmt/status" -UseBasicParsing
    $res = $r.RawContent
        if ($res.contains("OK")) {
            Show-Icon "success"
            Write-Host "REST-Schnittstelle liefert Status: OK auf $portmgmt"
        }
        else {
            Show-Icon "error"
            Write-Host "REST-Schnittstelle liefert Status: $res"
        }       
    }
    catch{
        Show-Icon "error"
        Write-Warning "REST-Schnittstelle konnte nicht angesprochen werden"
    } 
}


function CheckDocPortal {
	$ConnectPfad = Get-ItemProperty -Path "Registry::HKCU\Software\CompuGROUP\DocPortal" -Name binary -ErrorAction SilentlyContinue
	
	$connect = $ConnectPfad.binary
	if ( !$connect ) {
		Show-Icon "error"
        Write-Host "Registry-Docportal Pfad wurde nicht gefunden: \HKEY_CURRENT_USER\Software\CompuGROUP\DocPortal -> binary"
		return
	}
	
	if ( $connect -Match "mslib_local" ) {
		Show-Icon "success"
	    Write-Host "Registry-Docportal Pfad: $connect";
	}
	else {
		Show-Icon "error"
        Write-Host "Registry-Docportal Pfad ist falsch: $connect   -> Im Regisrtry (\HKEY_CURRENT_USER\Software\CompuGROUP\DocPortal) anpassen"
	}
}

function GetMedistarPath {

    # Prüfen ob Umgebungsvariable gesetzt worden ist
	$Medistar_Path = $ENV:medistardir
    if ( ! $Medistar_Path ) {
        Show-Icon "error"
        Write-Warning "MEDISTAR-Umgebungsvariable wurde nicht gesetzt"
        return #$false
    }
	
	# Prüfen, ob Pfad existiert
    if (! (Test-Path -path $Medistar_Path)) {
        Show-Icon "error"
        Write-Warning "MEDISTAR-Installationspfad existiert nicht: $Medistar_Path"
        return #$false
    }
	
    if ($Medistar_Path.Substring(0, 2) -eq "\\"){
        Show-Icon "error"
        Write-Warning "Den UNC-Pfad bitte korrigieren: $Medistar_Path"
        pause
        exit
    }else{
        Show-Icon "success"
        Write-Host "Medistar-Umgebungsvariable gefunden: $Medistar_Path"
    }
}

function CheckMsNet {	
	# Prüfen ob MSNETIN Verzeichnis vorhanden ist
	$msnetin = "MSNETIN"
	$fpath = Join-Path -path $ms_path -ChildPath $msnetin
	if (Test-Path -path $fpath ) {
		Show-Icon "success"
        Write-Host "Verzeichnis ist vorhanden: $fpath"
	}
	else {
		Show-Icon "error"
        Write-Host "Verzeichnis ist nicht vorhanden: $fpath  -> Bitte neu erstellen"
	}
	
    # Prüfen ob MSNETOU Verzeichnis vorhanden ist
	$msnetout = "MSNETOUT"
	$f2path = Join-Path -path $ms_path -ChildPath $msnetout
	if (Test-Path -path $f2path ) {
		Show-Icon "success"
        Write-Host "Verzeichnis ist vorhanden: $f2path"
	}
	else {
		Show-Icon "error"
        Write-Host "Verzeichnis ist nicht vorhanden: $f2path  -> Bitte neu erstellen"
	}
}

function Checkcardverification {
    $fehlerlog = "KIM\KIM_Clientmodul\logs\cm.fehler.log"
    $fpath = Join-Path -path $ms_path -ChildPath $fehlerlog

    #Prüfen, ob Datei existiert
    if (! (Test-Path -path $fpath)) {
        Show-Icon "error"
        Write-Host "Kein 'cm.fehler.log' gefunden: $fpath"
        return
    }

    #Wenn "has PinStatus 'VERIFIED'" im Log vorkommt, ist der PinStatus verifiziert
    if ((Get-Content $fpath) -Match "has PinStatus 'VERIFIED'") {
        Show-Icon "success"
        Write-Host "PinStatus ist verifiziert"
    }
    else {
        Show-Icon "error"
        Write-Warning "PinStatus ist nicht verifiziert"
    }

}

function MailCheck {
    $certificates_to_check= "KIM\KIM_Assist\data\kim\data\*.p12"
    $fpath = Join-Path -path $ms_path -ChildPath $certificates_to_check

    #Prüfen, ob Datei existiert
    if (! (Test-Path -path $fpath)) {
        Show-Icon "error"
        Write-Host "Keine *.p12-Datei in '$fpath'"
        return
    }

    $certificates_to_check2 = "KIM\KIM_Assist\data\kim\data\"
    $fpath = Join-Path -path $ms_path -ChildPath $certificates_to_check2

    Get-ChildItem -Path $fpath -Filter *.p12 -File -Name | ForEach-Object {
        $mailAdress = [System.IO.Path]::GetFileNameWithoutExtension($_)
        Show-Icon "success"
        Write-Host "KIM-Addresse: '$mailAdress'"
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

function ServicecheckKBV {
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
    #$komLeClientManagementPort = $connect.GeneralConfiguration.komLeClientManagementPort
    if ($pop3Port -and $smtpPort) {
        Show-Icon "success"
        Write-Host " - Ports (POP3, SMTP): $pop3Port, $smtpPort"
    }
    else {
        Show-Icon "error"
        Write-Warning " - Es sind nicht alle Ports (POP3 & SMTP) angegeben"
    }

    # Zertifikatspfad
    $certPath = $connect.GeneralConfiguration.ldapsCertificatePath

    if($certPath) {
        Show-Icon "success"
        Write-Host "Zertifikatspfad: $certPath"
    }
    else{
        Show-Icon "error"
        Write-Warning "INFO: Es ist kein Zertifikat hinterlegt"
    }

    # Zertifikatskennwort
    $certpass = $connect.GeneralConfiguration.ldapsCertificatePassword

    if($certpass){
        Show-Icon "success"
        Write-Host "verschluesseltes Zertifikatskennwort: $certpass"
    }
    else{
        Show-Icon "error"
        Write-Warning "INFO: Es ist kein Zertifikatskennwort hinterlegt"
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

    Copy-Item -Path $sys_path -Destination $ms_path"\sysconf.txt"
    
    #sysconf auslesen
    $content = Get-Content $ms_path"\sysconf.txt" | Where-Object {$_ -like "*MS4 = "+$ms_path+"\para*"}

    #Prüfen ob der UNC-Pfad hinterlegt ist
    #$ms4 = Get-Content $content

    if ($content -like "*MS4 = "+$ms_path+"\para*"){
        Show-Icon "success"
        Write-Host "lokaler Pfad in der sysconf.s"
    }
    else {
        Show-Icon "error"
        Write-Warning "UNC-Pfad, bitte korrigieren in den lokalen Pfad"
    }
}

function admin {
        
    $role = whoami /groups /fo csv | convertfrom-csv | where-object { $_.SID -eq "S-1-5-32-544" }

    if ($role -like "*Administratoren*"){
        Show-Icon "success"
        Write-Host "Nutzer hat administrative Rechte"
    }else{
        Show-Icon "error"
        Write-Warning "Nutzer hat keine administrativen Rechte"
    }
}

function dbms {

    $exe = "\prg4\m42t.exe"    
    $version = (Get-Item (Join-Path -path $ms_path -Childpath $exe)).VersionInfo.ProductVersion

    if ($version -ge $ms_version){
        Show-Icon "success"
        Write-Host "Medistar ist aktuell: $version"
    }else{
        Show-Icon "error"
        Write-Warning "Medistar ist nicht aktuell: $version"
    }
}

function connect {

    $connect = "\prg4\connect\connect.exe"
    $cVersion = (Get-item (Join-Path -path $ms_path -Childpath $connect)).VersionInfo.ProductVersion

    if ($cVersion -ge $connect_version){
        Show-Icon "success"
        Write-Host "Connect Version: $cVersion"
    }else{
        Show-Icon "error"
        Write-Warning "Connect ist nicht aktuell: $cVersion"
    }
}

function FWCheck{

    $com = powershell Get-NetFirewallRule -Displayname *KIM | select DisplayName, Direction, Enabled, PrimaryStatus

    if (! $com){
        Show-Icon "error"
        Write-Warning "keine Windows-Firewall-Regeln gesetzt"

    }else{
        Show-Icon "success"
        Write-Host "Windows-Firewall-Regeln wurden gesetzt"
    }

}

function routing{

    $route = Get-CimInstance -Class Win32_IP4PersistedRouteTable -Filter 'Caption LIKE "100.10%"' | Format-Table -Property Name,Mask,NextHop -AutoSize

    if (!$route){
        Show-Icon "error"
        Write-Warning "Keine Routen gefunden, IEGK pruefen"
    }else{
        Show-Icon "success"
        Write-Host "Routen gefunden:" 
        $route
    }
}


function multiprov{

    $sqlQuery = @"
    COLUMN CM_ADRESSE   format a16 HEADING "CM Adresse"
    COLUMN USERNAME     format a40 HEADING "KIM Adresse"
    COLUMN POP3         format 9999 HEADING "POP3"
    COLUMN SMTP         format 9999 HEADING "SMTP"

    SELECT  COMM_KIM_ACCOUNT_CM.CM_ADRESSE
    ,COMM_ACCOUNT.USERNAME
    ,COMM_KIM_ACCOUNT_CM.POP3
    ,COMM_KIM_ACCOUNT_CM.SMTP
    FROM COMM_KIM_ACCOUNT_CM
    INNER JOIN COMM_ACCOUNT
    ON COMM_ACCOUNT.KEYID = COMM_KIM_ACCOUNT_CM.COMM_ACCOUNT;
"@

$sqlQuery | sqlplus msuser/msuser1234@medistar
}

function timeout{
    $fehlerlog = "\KIM\KIM_Assist\data\kim\log\debug.log"
    $fpath = Join-Path -path $ms_path -ChildPath $fehlerlog

    #Prüfen, ob Datei existiert
    if (! (Test-Path -path $fpath)) {
        Show-Icon "error"
        Write-Host "Kein 'debug.log' gefunden: $fpath"
        return
    }

    #Wenn "has PinStatus 'VERIFIED'" im Log vorkommt, vermutlich Firewall im Einsatz
    if ((Get-Content $fpath) -Match "failed: Connection timed out") {
        Show-Icon "error"
        Write-Warning "Time-Out in der LOG-Datei gefunden!"
        Show-Icon "error"
        Write-Warning "$fpath"
    }
    else {
        Show-Icon "success"
        Write-Host "kein Time-Out in der LOG-Datei gefunden"
    }
}

function simulation{

    $inp = Read-Host -Prompt "IP Adresse des Konnektors: "

    if($inp -eq ""){
        return
    }else{
        nslookup accounts.tm.kim.telematik $inp
    }

    tracert 100.102.7.140
}

function internFail{
    $fehlerlog = "KIM\KIM_Clientmodul\logs\cm.fehler.log"
    $fpath = Join-Path -path $ms_path -ChildPath $fehlerlog

    #Prüfen, ob Datei existiert
    if (! (Test-Path -path $fpath)) {
        Show-Icon "error"
        Write-Warning "Kein 'cm.fehler.log' gefunden: $fpath"
        return
    }

    #Wenn "interner Fehler'" im Log vorkommt, Kartenleser neustarten
    if ((Get-Content $fpath) -Match "KonnektorExeption: Interner Fehler") {
        Show-Icon "error"
        Write-Warning "KonnektorExeption: Interner Fehler in der LOG Datei gefunden, bitte Kartenleser neustarten"
        Show-Icon "error"
        Write-Warning "$fpath"
    }
    else {
        Show-Icon "success"
        Write-Host "keine Meldung gefunden"
    }

}

function connectLOGservice{
    $fehlerlog = "\connect\log\Communicator.log"
    $fpath = Join-Path -path $appdata -ChildPath $fehlerlog

    #Prüfen, ob Datei existiert
    if (! (Test-Path -path $fpath)) {
        Show-Icon "error"
        Write-Warning "keine 'communicator.log': $fpath"
        return
    }

    #Wenn "Sending of message failed Couldn't connect to host, port: MEDISTAR, 8465; timeout -1" im Log vorkommt
    if ((Get-Content $fpath) -Match "Sending of message failed Couldn't connect to host, port") {
        Show-Icon "error"
        Write-Warning "'Couldn't connect to host, port' in der LOG Datei gefunden, bitte Dienst und Ports pruefen!"
        Show-Icon "error"
        Write-Warning "$fpath"
    }
    else {
        Show-Icon "success"
        Write-Host "keine Meldung gefunden"
    }

}

function mtaFail{
    $fehlerlog = "KIM\KIM_Clientmodul\logs\cm.fehler.log"
    $fpath = Join-Path -path $ms_path -ChildPath $fehlerlog

    #Prüfen, ob Datei existiert
    if (! (Test-Path -path $fpath)) {
        Show-Icon "error"
        Write-Warning "Kein 'cm.fehler.log' gefunden: $fpath"
        return
    }

    #Wenn "Error in MTA'" im Log vorkommt = verloren, ggfs auf Virenscanner prüfen
    if ((Get-Content $fpath) -Match "error in MTA Connection") {
        Show-Icon "error"
        Write-Warning "'Error in MTA Connection' gefunden, ggfs auf Virenscanner achten"
        Show-Icon "error"
        Write-Warning "$fpath"
    }
    else {
        Show-Icon "success"
        Write-Host "keine Meldung gefunden"
    }

}

function keystoreFail{
    $fehlerlog = "KIM\KIM_Clientmodul\logs\cm.fehler.log"
    $fpath = Join-Path -path $ms_path -ChildPath $fehlerlog

    #Prüfen, ob Datei existiert
    if (! (Test-Path -path $fpath)) {
        Show-Icon "error"
        Write-Warning "Kein 'cm.fehler.log' gefunden: $fpath"
        return
    }

    #Wenn "Keystorepassword was incorrect" im Log vorkommt = über IKIM ein neues anlegen
    if ((Get-Content $fpath) -Match "keystore password was incorrect") {
        Show-Icon "error"
        Write-Warning "'keystore password was incorrect' gefunden, ggfs neues ueber IKIM anlegen"
        Show-Icon "error"
        Write-Warning "$fpath"
    }
    else {
        Show-Icon "success"
        Write-Host "keine Meldung gefunden"
    }

}

#Script Version
#updateS

#Banner anzeigen:
PrintBanner
Write-Host ""
Write-Host ""

$inpt = Read-Host -Prompt "CGM KIM [c] oder kv.dox [k] ?"

if ($inpt -eq "s"){
    Start-Process "https://www.memski.org/snake.html"
    return
}

if ($inpt -eq "a"){
    
    Write-Host ""
    Write-Host ""
    Show-Icon "error"
    Write-Host "//TODO -> REST API requests einbauen"
    Write-Host ""
    Pause
    return
}

if ($inpt -eq "k"){
   
    Write-Host ""
    Write-Host "  Achtung!"
    Write-Host ""
    Write-Host ""
    pause
        
    Write-Host ""
    Write-Host "  Es geht gleich los!"
    Write-Host ""
    Write-Host ""
    pause

    Write-Host ""
    Write-Host "  Alle anschnallen!"
    Write-Host ""
    Write-Host ""
    pause

    #Dienst
    Write-Host ""
    Write-Host "  Status Windowsdienst"
    Write-Host "  ---------------------------------"
    foreach ($service in $services_to_check_kbv) {
    ServicecheckKBV $service
    }
    Write-Host ""
     
    #Ports checken:
    Write-Host ""
    Write-Host "  Checken der Ports"
    Write-Host "  ---------------------------------"
    foreach ($port in $ports_to_check_kbv) {
    Portcheck $port
    }
    Write-Host ""

    #Connect Version:
    Write-Host ""
    Write-Host "  Connect Version"
    Write-Host "  ---------------------------------"
    connect
    Write-Host ""
     
    #Plugin checken:
    Write-Host ""
    Write-Host "  Connect KIM-Plugin Konfiguration"
    Write-Host "  ---------------------------------"
    Plugin
    Write-Host ""
     
    Write-Host ""
    Write-Host " Info vom Autor"
    Write-Host "  ---------------------------------"
    Show-Icon "success"
    Write-Host "Mehr ist erstmal nicht interessant..."
    Write-Host ""
    Write-Host ""
    Pause
    Return
}

if ($inpt -eq "c"){
     
    #Umgebungsvariable
    Write-Host ""
    Write-Host "  Umgebungsvariable/Medistar-Pfad"
    Write-Host "  ---------------------------------"
    GetMedistarPath
    Write-Host ""
    pause
             
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
          
    #KIM Einrichtungsassist jar
    Write-Host ""
    Write-Host "  KIM Einrichtungsassist.jar"
    Write-Host "  ---------------------------------"
    KimAssistInstallationCheck
    Write-Host ""
          
    #aktuelle CM Version
    Write-Host ""
    Write-Host "  KIM CM Version"
    Write-Host "  ---------------------------------"
    KimClientmodulInstallationCheck
    Write-Host ""
         
    #Check MsNet-Ordner
    Write-Host ""
    Write-Host "  Check MsNet-Ordner"
    Write-Host "  ---------------------------------"
    CheckMsNet
    Write-Host ""
          
    #Ports checken:
    Write-Host ""
    Write-Host "  Checken der Ports"
    Write-Host "  ---------------------------------"
    foreach ($port in $ports_to_check_cgm) {
        Portcheck $port
    }
    Write-Host ""
          
    #sysconf.s checken
    Write-Host ""
    Write-Host "  Checken der sysconf"
    Write-Host "  ---------------------------------"
    sysconf
    Write-Host ""
             
    #routen
    Write-Host ""
    Write-Host "  Checken der Routen"
    Write-Host "  ---------------------------------"
    routing
    Write-Host ""
    
    #Services checken:
    Write-Host ""
    Write-Host "  Status Windowsdienst"
    Write-Host "  ---------------------------------"
    foreach ($service in $services_to_check_cgm) {
        Servicecheck $service
    }
    Write-Host ""

    #Connect Version:
    Write-Host ""
    Write-Host "  Connect Version"
    Write-Host "  ---------------------------------"
    connect
    Write-Host ""
          
    #Plugin checken:
    Write-Host ""
    Write-Host "  Connect KIM-Plugin Konfiguration"
    Write-Host "  ---------------------------------"
    Plugin
    Write-Host ""
          
    #REST-Check:
    Write-Host ""
    Write-Host "  REST-Check"
    Write-Host "  ---------------------------------"
    RESTCheck
    Write-Host ""
          
    #CG Java Pfad:
    Write-Host ""
    Write-Host "  CG Java"
    Write-Host "  ---------------------------------"
    MedistarJavaPathCheck
    Write-Host ""
          
    #Docportal Registry-Eintrag:
    Write-Host ""
    Write-Host "  Docportal Registry-Eintrag"
    Write-Host "  ---------------------------------"
    CheckDocPortal
    Write-Host ""
          
    #P12-Zertifikate:
    Write-Host ""
    Write-Host "  P12-Zertifikate"
    Write-Host "  ---------------------------------"
    MailCheck
    Write-Host ""
          
    #SMCB Pin Status aus der LOG:
    Write-Host ""
    Write-Host "  SMC-B Pin Status"
    Write-Host "  ---------------------------------"
    Checkcardverification
    Write-Host ""
          
    #Secret checken:
    Write-Host ""
    Write-Host "  Secret Datei"
    Write-Host "  ---------------------------------"
    Secret
    Write-Host ""

    #FW checken:
    Write-Host ""
    Write-Host "  Windows Firewall"
    Write-Host "  ---------------------------------"
    FWCHeck
    Write-Host ""

    #Multikonnektor checken:
    Write-Host ""
    Write-Host "  Multikonnektoranbindung"
    Write-Host "  ---------------------------------"
    multiprov
    Write-Host ""

    #cm.fehler.log - Interner Fehler:
    Write-Host ""
    Write-Host "  cm.fehler.log - Interner Fehler"
    Write-Host "  ---------------------------------"
    internFail
    Write-Host ""

    #cm.fehler.log - Interner Fehler:
    Write-Host ""
    Write-Host "  communicator.log - 'Couldn't connect to host, port'"
    Write-Host "  ---------------------------------"
    connectLOGservice
    Write-Host ""

    #cm.fehler.log - MTA Error:
    Write-Host ""
    Write-Host "  cm.fehler.log - 'Error in MTA Connection'"
    Write-Host "  ---------------------------------"
    mtaFail
    Write-Host ""

    #cm.fehler.log - keystore password was incorrect:
    Write-Host ""
    Write-Host "  cm.fehler.log - 'keystore password was incorrect'"
    Write-Host "  ---------------------------------"
    keystoreFail
    Write-Host ""
    
    #LOG Datei auf Time Out prüfen:
    Write-Host ""
    Write-Host "  debug.log - Time-Out Accountmanager"
    Write-Host "  ---------------------------------"
    timeout
    Write-Host ""

    #Verbindungssimulation:
    Write-Host ""
    Write-Host "  Verbindung Accountmanager"
    Write-Host "  ---------------------------------"
    simulation
    Write-Host ""
    Write-Host ""
    Write-Host ""

    pause
    Return
}