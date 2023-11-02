$headers=@{}

#Key fra https://SITENAME/account/api
$key = ""
$siteurl = 'https://SITENAME/api/v1'

$headers.Add("accept", "application/json")
$headers.Add("content-type", "application/json")
$headers.Add("authorization", "Bearer " + $key)

$Date = '{0:yyyy-MM-dd}' -f (Get-Date)
$PC = $env:COMPUTERNAME
$filename = "$PC-$Date.txt"

$folder = "C:\Logs"

$output = "$folder\SUCCES\$PC-$Date.txt"
$outputError = "$folder\ERROR\$PC-$Date.txt"
$outputFatal = "$folder\FATAL\$PC-$Date.txt"

#dårlig kode under
$fatalCount = 0
$errorCount = 0

#######
 # Tjek om mappe eksisterer, hvis ikke så opret.
 #
 # @param {string} Lokation - F.eks. om mappen C:\Logs eksisterer
######
Function Folder-CheckCreate()
{
    param(
    [Parameter (Mandatory = $true)] [String]$Lokation
    )  

    If (!(test-path $Lokation))
    {
        [void](New-Item -Path $Lokation -ItemType Directory)
    }
}

Folder-CheckCreate -Lokation "$folder"
Folder-CheckCreate -Lokation "$folder\SUCCES"
Folder-CheckCreate -Lokation "$folder\ERROR"
Folder-CheckCreate -Lokation "$folder\FATAL"

#######
 # Opretter et nyt, eller tilføjer til eksisterende log
 #
 # @param {string} Level - INFO, WARN, ERROR, FATAL, DEBUG eller SUCCES kan bruges. Den viser dette i selve log filen.
 # @param {string} Message - Beskeden i log filen
 # @param {string} logfile - Logfilens output placering, f.eks. C:\logs (Se variablen $output på linje 14)
######
Function Write-Log {
    [CmdletBinding()]
    Param(
    [Parameter(Mandatory=$False)]
    [ValidateSet("INFO","WARN","ERROR","FATAL","DEBUG", "SUCCES")]
    [String]
    $Level = "INFO",

    [Parameter(Mandatory=$True)]
    [string]
    $Message,

    [Parameter(Mandatory=$False)]
    [string]
    $logfile
    )

    if ($Level -eq "FATAL") {
        $script:fatalCount++
    }

    if ($Level -eq "ERROR") {
        $script:errorCount++
    }

    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $Line = "$Stamp $Level $Message"

    If($logfile) {
        Add-Content $logfile -Value $Line
    }
    Else {
        Write-Output $Line
    }
}

#######
 # Bruges til at teste forbindelsen til API
 # 
 # Returner enten true eller false.
######
function testConnection
{
    try {
        $testSite = Invoke-WebRequest -Uri "$siteurl/statuslabels" -Method GET -Headers $headers
        if ($testSite.statuscode -eq '200') {
            Write-Log -Level SUCCES -Message "[SNIPEIT]-[testConnection] SUCCESFULLY CONNECTED TO API" -logfile $output
            return $true
        }
    } catch {
        Write-Log -Level FATAL -Message "[SNIPEIT]-[testConnection] CAN'T CONNECT TO API" -logfile $output
        return $false
    }
}

#######
 # Returner et object med data på computer fra Snipe IT
 #
 # @param {string} type - Valgte type at søge efter bytag eller byserial (ps://snipe-it.readme.io/reference/hardware-by-asset-tag)
 # @param {string} pcCode - Søg efter enten PC-123456 eller serie nr
######
function getAssetinfo
{
    param(
    [Parameter (Mandatory = $true)] [String]$type,
    [Parameter (Mandatory = $true)] [String]$pcCode
    )
    try {
        $getData = Invoke-RestMethod -Uri "$siteurl/hardware/$type/$pcCode" -Method GET -Headers $headers
        if ($type -eq "byserial") {
            Write-Log -Message "[SNIPEIT]-[getAssetinfo] - GETTING PC DATA FROM SERIAL: $pcCode" -logfile $output
            $getData = $getData.rows
        }
    
        if ($getData.status -eq 'error') 
        {
            return $getData.messages
        } else {
            foreach ($data in $getData) {
                if ($data.deleted_at -eq $null)  {
                    return $data
                    break
                }
            }
        }
    } catch {
        $error = $_.ScriptStackTrace
        Write-Log -Level ERROR -Message "[SNIPEIT]-[getAssetinfo] - UNEXPECTED ERROR: $error" -logfile $output
        Write-Log -Level ERROR -Message "[SNIPEIT]-[getAssetinfo] - DESCRIPTION: $_" -logfile $output
        return
    }
}

#######
 # Returner et object med PC data fra Snipe IT
 #
 # @param {string} pcCode - Søg efter PC-123456
 # @param {string} pserial - Søg efter serie nr f.eks. 5CGH349SJA
######
function getAssetData
{
    param(
    [Parameter (Mandatory = $true)] [String]$pcCode,
    [Parameter (Mandatory = $true)] [String]$pserial
    )

    try {
        $snipeData = @()
        $byAsset = getAssetInfo -type "bytag" -pcCode "asd"
        $bySerial = getAssetInfo -type "byserial" -pcCode $pserial

        $notFound = "Asset does not exist." #Output fra byAsset eller bySerial

        if ($byAsset -ne $notFound) {
            Write-Log -Level INFO -Message "[SNIPEIT]-[getAssetData] - FOUND PC BY ASSET TAG: $pcCode" -logfile $output
            $snipeData = $byAsset
        } elseif ($bySerial -ne $notFound) {
            Write-Log -Level INFO -Message "[SNIPEIT]-[getAssetData] - COULDN'T FIND PC BY ASSET TAG, FOUND BY SERIAL: $pserial" -logfile $output
            $snipeData = $bySerial
        } else {
            Write-Log -Level FATAL -Message "[SNIPEIT]-[getAssetData] - COULDN'T FIND PC BY ASSET TAG OR SERIAL" -logfile $output
            return
        }

        $properties = @{
            assetTag = $snipeData.asset_tag
            serial = $snipeData.serial
            model =  $snipeData.manufacturer.name + " " +$snipeData.model.name
            ram = $snipeData.custom_fields.RAM.value
            cpu = $snipeData.custom_fields.CPU.value
            storage = $snipeData.custom_fields.Storage.value
        }

        $contentToString = $properties | Out-String
        Write-Log -Level SUCCES -Message "[SNIPEIT]-[getAssetData] - DATA COLLECTED FROM SNIPEIT" -logfile $output

        return New-Object psobject -Property $properties
    } catch {
        $error = $_.ScriptStackTrace
        Write-Log -Level ERROR -Message "[SNIPEIT]-[getAssetData] - UNEXPECTED ERROR: $error" -logfile $output
        Write-Log -Level ERROR -Message "[SNIPEIT]-[getAssetData] - DESCRIPTION: $_" -logfile $output
        return        
    }
}


#######
 # Returner en integer med værdi tættest på
 #
 # @param {string} diskSize - Størrelse af disk størrelse.
######
function getdisk
{
    param(
    [Parameter (Mandatory = $true)] [String]$diskSize
    )

    $availableSizes = 128,256,512,1000,2000
    $oldval = $diskSize - $availableSizes[0]
    $Final = $availableSizes[0]

    if($oldval -lt 0){$oldval = $oldval * -1}

    $availableSizes | %{$val = $diskSize - $_

    if($val -lt 0 ){$val = $val * -1}
    if ($val -lt $oldval){

    $oldval = $val
    $Final = $_} }

    return $Final.ToString() + " GB"
}

#######
 # Returner et object med nuværende PC data
 #
######
function getLocaldata
{
    try {
        $data = Get-CimInstance -ClassName Win32_ComputerSystem
        $cpuData = gwmi win32_Processor
        $cDrivespace = Get-Volume -DriveLetter C | Select -ExpandProperty Size

        $heimdalAdmins = Get-LocalGroupMember -Group "Heimdal" -ErrorAction SilentlyContinue | Select -ExpandProperty Name

        if ($heimdalAdmins -eq $null) {
            $dom = "null"
            $usr = "null"
            $fullname = "null"
            Write-Log -Level WARN -Message "[SNIPEIT]-[getLocaldata] - HEIMDAL NOT FOUND, USER NULL" -logfile $output
        } else {
            $dom = $heimdalAdmins.Split('\')[0]
            $usr = $heimdalAdmins.Split('\')[-1]
            $fullname = ([adsi]"WinNT://$dom/$usr,user") | Select-Object -ExpandProperty fullname
            Write-Log -Level SUCCES -Message "[SNIPEIT]-[getLocaldata] - HEIMDAL USER FOUND: $usr" -logfile $output
        }

        $ram = (Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum /1gb
        $ramtoStr = [string]$ram

        $properties = @{
            assetTag = $env:COMPUTERNAME
            serial = Get-WmiObject win32_bios | Select-Object -ExpandProperty Serialnumber
            manufacturer =  $data.Model
            model =  $data.Model.replace($data.Manufacturer + " ", "")
            ram = "$ramtoStr GB"
            cpu = $cpuData.Name
            realdisk = $cDrivespace / 1GB
            storage = getdisk -diskSize ($cDrivespace / 1GB)
            user = $usr
            domain = $dom
            fullname = $fullname
        }

    $contentToString = $properties | Out-String
    Write-Log -Level SUCCES -Message "[SNIPEIT]-[getLocaldata] - DATA COLLECTED FROM LOCAL PC" -logfile $output
    return New-Object psobject -Property $properties
    } catch {
        $error = $_.ScriptStackTrace
        Write-Log -Level ERROR -Message "[SNIPEIT]-[getLocaldata] - UNEXPECTED ERROR: $error" -logfile $output
        Write-Log -Level ERROR -Message "[SNIPEIT]-[getLocaldata] - DESCRIPTION: $_" -logfile $output
        return
    }
}

#######
 # Bruges til at sammenligne data, dog ikke aktuelt pt.
 #
 # @param {object} localData - Lokal data fra computeren fra i dette objekt.
 # @param {object} snipeData - SnipeIT data i dette objekt.
######
function matchData
{
    param(
    [Parameter (Mandatory = $true)] [Object[]]$localData,
    [Parameter (Mandatory = $true)] [Object[]]$snipeData
    )

    Compare-Object -ReferenceObject $localData -DifferenceObject $snipeData -Property assetTag, serial, ram, cpu, storage
}

#######
 # For at få computer ID fra SnipeIT, det er lidt bøvlet. Derfor nedenstående metode
 #
 # @param {String} PC - F.eks. HP EliteBook 840 G7
######
function getModel()
{
    param(
    [Parameter (Mandatory = $true)] [String]$PC
    )

    try {
        $models = Invoke-RestMethod -Uri "$siteurl/models" -Method GET -Headers $headers
        
        foreach ($model in $models.rows)
        {
            $name = $model.name
            
            #if ("*EliteBook 840 G7 Notebook PC*" -like $name) {
            if ("*$PC*" -match $name) {
                $properties = @{
                    id = $model.id
                    name = $name
                    manufacturer_id = $model.manufacturer.id
                    manufacturer_name= $model.manufacturer.name
                }
                $contentToString = $properties | Out-String
                Write-Log -Level SUCCES -Message "[SNIPEIT]-[getModel] - FOUND MATCHING MODEL: $name" -logfile $output
                #Write-Log -Message "[SNIPEIT]-[getModel] - DATA COLLECTED FROM SNIPEIT: $contentToString" -logfile $output
                return New-Object psobject -Property $properties
            }
        }
    } catch {
        $error = $_.ScriptStackTrace
        Write-Log -Level ERROR -Message "[SNIPEIT]-[getModel] - UNEXPECTED ERROR: $error" -logfile $output
        Write-Log -Level ERROR -Message "[SNIPEIT]-[getModel] - DESCRIPTION: $_" -logfile $output
        return
    }
}

#######
 # Hiver bruger data udfra SnipeIT som er nødvendigt til oprettelse eller opdatering af PC.
 #
 # @param {String} username - Kan eksempelvis være "torben" - dette kan fåes igennem whoami eller Heimdal
######
function getSnipeITuser()
{
    param(
    [Parameter (Mandatory = $true)] [String]$username
    )

    try {
        $getSearch = Invoke-RestMethod -Uri "$siteurl/users?username=$username&limit=50&offset=0&sort=created_at&order=desc&deleted=false&all=false" -Method GET -Headers $headers

        if ($username -eq "null") {
            $properties = @{
                userid = "null"
                name = "null"
                username = "null"
                department_id = "null"
                department_name = "null"
                location_id = "null"
                location_name = "null"
                company = "null"
            }

            Write-Log -Level WARN -Message "[SNIPEIT]-[getSnipeITuser] - DIDN'T FOUND SNIPEIT USER" -logfile $output
            #$contentToString = $properties | Out-String
            #Write-Log -Message "[SNIPEIT]-[getSnipeITuser] - DATA COLLECTED FROM SNIPEIT: $contentToString" -logfile $output
            return New-Object psobject -Property $properties
        }
        

        foreach ($field in $getSearch) {
            if ($field.total -eq 1) {
                $properties = @{
                    userid = $field.rows.id
                    name = $field.rows.name
                    username = $field.rows.username
                    department_id = $field.rows.department.id
                    department_name = $field.rows.department.name
                    location_id = $field.rows.location.id
                    location_name = $field.rows.location.name
                    company = $field.rows.company.id
                }
                $name = $field.rows.name
                Write-Log  -Level SUCCES -Message "[SNIPEIT]-[getSnipeITuser] - FOUND SNIPEIT USER: $name" -logfile $output
                return New-Object psobject -Property $properties      
            }
        }
    } catch {
        $error = $_.ScriptStackTrace
        Write-Log -Level ERROR -Message "[SNIPEIT]-[getModel] - UNEXPECTED ERROR: $error" -logfile $output
        Write-Log -Level ERROR -Message "[SNIPEIT]-[getModel] - DESCRIPTION: $_" -logfile $output
        return
    }
}

#######
 # Oprettelse af komplet ny device i SnipeIT
 #
 # @param {object} localData - Alt lokal data fra computeren i dette objekt.
######
#https://github.com/snipe/snipe-it/issues/10755#issuecomment-1055920645
function createPC()
{
    param(
    [Parameter (Mandatory = $true)] [Object]$localData
    )

    Write-Log -Level SUCCES -Message "[=================== CREATING PC STARTED ===================] " -logfile $output

    try {
        $localData = getLocaldata
        $pcData = getModel -PC $localData.model
        $snipeitUser = getSnipeITuser -username $localData.user

        #status_id
        #1 = Pending
        #2 = Servicedesk in-house
        #3 = Archived
        #4 = I drift
        #$siteurl/statuslabels
        $create = @{
            "archived" = $false
             "warranty_months" = $null
             "depreciate" = $false
             "supplier_id" = $null
             "requestable" = $false
             "rtd_location_id" = $snipeitUser.location_id
             "last_audit_date" = "null"
             "location_id"  = $snipeitUser.location_id
             #"asset_tag" = $localData.assetTag
             "asset_tag" = $localData.assetTag
             "status_id" = "4"
             "model_id" = $pcData.id
             "categories" = "2"
             "serial" = $localData.serial
             "_snipeit_ram_3" = $localData.ram
             "_snipeit_cpu_4" = $localData.cpu
             "_snipeit_storage_8" = $localData.storage
             "company_id" = "2" #@{id=2; name=COMPANY NAME}
             "assigned_user" = $snipeitUser.userid
        } | ConvertTo-Json

        $send = Invoke-RestMethod -Uri "$siteurl/hardware" -Method POST -Headers $headers -ContentType 'application/json' -Body $create
        $status = $send.status

        if ($status -eq "error") {
            Write-Log -Level FATAL -Message "[SNIPEIT]-[createPC] - FAILED TO CREATE PC: $status" -logfile $output
        } else {
            $contentToString = $create | Out-String
            Write-Log -Level SUCCES -Message "[SNIPEIT]-[createPC] - SUCCESFULLY CREATED PC $contentToString" -logfile $output
        }
    } catch {
        $error = $_.ScriptStackTrace
        Write-Log -Level ERROR -Message "[SNIPEIT]-[createPC] - UNEXPECTED ERROR: $error" -logfile $output
        Write-Log -Level ERROR -Message "[SNIPEIT]-[createPC] - DESCRIPTION: $_" -logfile $output
        return
    }
    Write-Log -Level SUCCES -Message "[=================== CREATING PC END ===================] " -logfile $output
}

#######
 # Opdatering af device i SnipeIT
 #
 # @param {object} localData - Alt lokal data fra computeren i dette objekt.
######
function updateData()
{
    param(
    [Parameter (Mandatory = $true)] [Object[]]$localData
    )

    Write-Log -Level SUCCES -Message "[=================== UPDATING PC STARTED ===================] " -logfile $output

    try {
        $localData = getLocaldata
        $pcData = getModel -PC $localData.model

        $snipeitUser = getSnipeITuser -username $localData.user
        $assetInfo = getAssetInfo -type "bytag" -pcCode $localData.assetTag
    
        $update = @{
             "asset_tag" = $localData.assetTag
             "archived" = $false
             "rtd_location_id" = $snipeitUser.location_id
             "location_id"  = $snipeitUser.location_id
             "status_id" = "4"
             "model_id" = $pcData.id
             "serial" = $localData.serial
             "_snipeit_ram_3" = $localData.ram
             "_snipeit_cpu_4" = $localData.cpu
             "_snipeit_storage_8" = $localData.storage
             "company_id" = "2" #@{id=2; name=COMPANY NAME}
             #"assigned_user" = $snipeitUser.userid
        } | ConvertTo-Json

        $pcID = $assetInfo.rows.id

        #https://snipe-it.readme.io/reference/hardware-checkout
        #SnipeIT kan ikke skelne, så hvis man ikke tjekke vil der stå "Checkout XX" i oversigten
        if ($snipeitUser.userid -ne $assetInfo.rows.assigned_to.id) {
            $update = @{
             "checkout_to_type" = "user"
             "assigned_user" = $snipeitUser.userid
            } | ConvertTo-Json 
            $username = $snipeitUser.username

            $checkOut = Invoke-RestMethod -Uri "$siteurl/hardware/$pcID/checkout" -Method POST -Headers $headers -ContentType 'application/json' -Body $update
            Write-Log -Level INFO -Message "[SNIPEIT]-[updateData] - CHECKED OUT TO USER $username" -logfile $output
        }

        $send = Invoke-RestMethod -Uri "$siteurl/hardware/$pcID" -Method PATCH -Headers $headers -ContentType 'application/json' -Body $update

        if ($send.status -eq "error") {
            Write-Log -Level FATAL -Message "[SNIPEIT]-[updateData] - FAILED TO UPDATE DATA" -logfile $output
        } else {
            Write-Log -Level SUCCES -Message "[SNIPEIT]-[updateData] - SUCCESFULLY UPDATED DATA" -logfile $output
        }
    } catch {
        $error = $_.ScriptStackTrace
        Write-Log -Level ERROR -Message "[SNIPEIT]-[updateData] - UNEXPECTED ERROR: $error" -logfile $output
        Write-Log -Level ERROR -Message "[SNIPEIT]-[updateData] - DESCRIPTION: $_" -logfile $output
        return
    }
    Write-Log -Level SUCCES -Message "[=================== UPDATING PC END ===================] " -logfile $output
}

function main 
{
    try {
        if (testConnection) {
            $localData = getLocaldata
            $snipeitData = getAssetData -pcCode $localData.assetTag -pserial $localData.serial

            if ($snipeitData -ne $null) {
                updateData -localData $localData
            } else {
                createPC -localData $localData
            }
        } else {
            Write-Log -Level FATAL -Message "[SNIPEIT]-[main] - NO CONNECTION TO API, QUIT" -logfile $output
        }
    } catch {
        $error = $_.ScriptStackTrace
        Write-Log -Level ERROR -Message "[SNIPEIT]-[main] - UNEXPECTED ERROR: $error" -logfile $output
        Write-Log -Level ERROR -Message "[SNIPEIT]-[main] - DESCRIPTION: $_" -logfile $output        
    }

    if ($errorCount -ge 1) {
        try {
            Move-Item -Path $output -Destination $outputError -Force
        } catch {}
    }

    if ($fatalCount -ge 1) {
        try {
            Move-Item -Path $output -Destination $outputFatal
        } catch {}
    }
}


main
