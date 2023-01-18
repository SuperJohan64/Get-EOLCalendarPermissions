# Connects to Exchange Online.
Try {Connect-ExchangeOnline -ErrorAction Stop}
Catch {Exit}

# Launches an open file dialog window from .Net to select the CSV File.
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.initialDirectory = $PSScriptRoot
$OpenFileDialog.filter = "CSV (*.csv)| *.csv"
$OpenFileDialog.ShowDialog() | Out-Null
$OpenFileDialog.filename
$CsvInput = $OpenFileDialog.filename

# Imports the data from the CSV file.
Try {$CsvData = Import-Csv -Path $CsvInput -ErrorAction Stop}
Catch {Exit}

# Variables used for output file.
$Timestamp = Get-Date -Format MMddyyyyHHmmss
$OutputFileName = ($MyInvocation.MyCommand.Name).Replace(".ps1","") + "-$Timestamp"
$OutputFilePath = "$PSScriptRoot\$OutputFileName.txt"

$CalendarPermissions = @()

foreach ($User in $CsvData) {
    $CalendarPath = $User.UserPrincipalName + ":\Calendar"
    $CalendarDetails = Get-MailboxFolderPermission -Identity $CalendarPath
    $CalendarPermissions += New-Object -TypeName psobject -Property @{
        MailboxOwnersName = $User.Name
        MailboxOwnersEmail = $User.EmailAddress
        MailboxOwnersUserPrincipalName = $User.UserPrincipalName
        Identity = ($CalendarDetails | Select-Object Identity -ExpandProperty Identity)
        FolderName = ($CalendarDetails | Select-Object FolderName -ExpandProperty FolderName)
        User = ($CalendarDetails | Select-Object User -ExpandProperty User)
        AccessRights = ($CalendarDetails | Select-Object AccessRights -ExpandProperty AccessRights)
        SharingPermissionFlags = ($CalendarDetails | Select-Object SharingPermissionFlags -ExpandProperty SharingPermissionFlags)
    }
}

$CalendarPermissions | `
    Select-Object MailboxOwnersName, MailboxOwnersEmail, MailboxOwnersUserPrincipalName, Identity, FolderName, User, AccessRights, SharingPermissionFlags | `
    Sort-Object MailboxOwnersEmail | `
    Out-File -FilePath $OutputFilePath

ii $OutputFilePath 
