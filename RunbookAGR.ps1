$AutomationAccountName = "aa-governance"
$AutomationResourceGroup = "rg-automation"
$ResourceGroup = 'FileStorageRG'
$SourceStorageAccountName = 'stlhrprivatesa'
$ShareName = 'scripts'
$Path = 'AzGovVizParallel.ps1'
$FilePath = "C:\Temp\AzGovVizParallel.ps1"
$OutputFolder = 'C:\Temp\Runbook_AGR'
$Output = "-NoJsonExport -Output '$OutputFolder'"
$Command = "$FilePath $Output"
$Exception = $false
$DstFile = "C:\Temp\RunBook_AGR.zip"
$Email = @{}
$Subject = "Azure Governance Reporting"
$Body = ""

# Ensures you do not inherit an AzContext in your runbook
Disable-AzContextAutosave -Scope Process | Out-Null

try {    
    # Connect to Azure with system-assigned managed identity
    $AzureContext = (Connect-AzAccount -Identity).context
}
catch {
    Write-Output "There is no system-assigned user identity. Aborting."; 
    Exit
}

# Set and store context
$AzureContext = Set-AzContext -SubscriptionName $AzureContext.Subscription -DefaultProfile $AzureContext

# Get SMTP information from the Automation Account variables
$SMTPUser = (Get-AzAutomationVariable -ResourceGroupName $AutomationResourceGroup -AutomationAccountName $AutomationAccountName -Name 'SMTPUser').Value
# $SMTPPass = Get-AutomationVariable -Name 'SMTPPass'
$SMTPServer = (Get-AzAutomationVariable -ResourceGroupName $AutomationResourceGroup -AutomationAccountName $AutomationAccountName -Name 'SMTPServer').Value
$SMTPPort = (Get-AzAutomationVariable -ResourceGroupName $AutomationResourceGroup -AutomationAccountName $AutomationAccountName -Name 'SMTPPort').Value
# $UseSsl = (Get-AzAutomationVariable -ResourceGroupName $AutomationResourceGroup -AutomationAccountName $AutomationAccountName -Name 'SMTPUseSsl').Value
$To = (Get-AzAutomationVariable -ResourceGroupName $AutomationResourceGroup -AutomationAccountName $AutomationAccountName -Name 'ToAddress').Value
# $Credential = (New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SMTPUser, (ConvertTo-SecureString -String $SMTPPass -AsPlainText -Force))

try {
    # Get Storage Account Key for the storage account containing the Azure Governance Reporting scripts
    Write-Output "Getting storage account key -> $ResourceGroup -> $SourceStorageAccountName"
    $SourceKey = Get-AzStorageAccountKey -ResourceGroupName $ResourceGroup -AccountName $SourceStorageAccountName

    # Get Storage Account
    Write-Output "Getting storage account context -> $SourceStorageAccountName"
    $StorageContext = New-AzStorageContext -StorageAccountName $SourceStorageAccountName -StorageAccountKey $SourceKey[0].Value

    # Remove $OutputFolder + $FilePath if they already exist
    # Create $OutputFolder
    if (Test-Path -Path $OutputFolder) {
        Remove-Item -Path $OutputFolder -Recurse -Force
    }
    [void](New-Item -Path $OutputFolder -ItemType Directory)
    if (Test-Path -Path $FilePath) {
        Remove-Item -Path $FilePath -Recurse -Force        
    }

    # Download script from file share in Storage Account
    Write-Output "Downloading file from the storage account -> $Sharename -> $Path -> C:\Temp"
    Get-AzStorageFileContent -ShareName $Sharename -Context $StorageContext -Path $Path -Destination "C:\Temp" -Force

    # Run download script
    Write-Output "Running downloaded script -> $Command"
    Invoke-Expression $Command
}
catch {
    $Exception = $true
    Write-Output $_.Exception
}

# Format Email properties
if ($Exception) {
    $Subject += " - Error"
    $Body += "`n`nAn error occurred while running this automated runbook. Please contact your support team for assistance.`n`n-Automation Team"
}
else {
    $Body += "`n`nAttached is a compressed file containing the documents generated from the Azure Governance Reporting tool.`n`n-Automation Team"
    $Email.Add('Attachments', $DstFile)

    # Test file and folder paths -> remove destination folder if exist
    Write-Verbose "Verifying source folder exist -> $OutputFolder"
    if (!(Test-Path -Path $OutputFolder)) {
        Write-Output "Source folder does not exist"
        Write-Output "Nothing to send"
        Write-Output "Stopping logging and terminating script"
        Stop-Transcript
        Exit
    }
    if (Test-Path -Path $DstFile) {
        Write-Output "Deleting destination file: $DstFile"
        Remove-Item -Path $DstFile -Force
    }

    # Compress files
    Write-Output "Compressing report files to: $DstFile"
    try {
        $Files = Get-ChildItem -Path $OutputFolder -Exclude 'JSON'
        Compress-Archive -Path $Files -DestinationPath $DstFile -ErrorAction Stop
    }
    catch {
        Write-Output "Unable to compress report files -> $OutputFolder -> $DstFile"
        Write-Output $_.Exception.Message
    }
}
$Email.Add('From', $SMTPUser)
$Email.Add('To', $To)
$Email.Add('Subject', $Subject)
$Email.Add('Body', $Body)
$Email.Add('SMTPServer', $SMTPServer)
$Email.Add('Port', $SMTPPort)
# $Email.Add('UseSsl', $UseSsl)
# $Email.Add('Credential', $Credential)

# Send email
Write-Output "Attempting to send the email"
try {
    Send-MailMessage @Email -ErrorAction Stop
    Write-Output "Email successfully sent"
} 
catch {
    Write-Output "Email unsuccessful"
    Write-Output $_.Exception.Message
}

# File clean up
if (Test-path -Path $OutputFolder) {
    Remove-Item -Path $OutputFolder -Recurse -Force
}
if (Test-path -Path $DstFile) {
    Remove-Item -Path $DstFile -Force
}