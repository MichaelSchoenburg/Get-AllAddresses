<#
.SYNOPSIS
    Get-AllAddresses
.DESCRIPTION
    PowerShell Script to read all addresses (including aliases) from all entities 
    (regular mailbox, shared mailbox, mail enabled public folder, contact, 
    unified group, distribution group, security group, ...) 
    there are in a microsoft 365 tenant.
.INPUTS
    None. You cannot pipe objects to Get-AllAddresses.
.OUTPUTS
    None/interactive.
.NOTES
    Author: Michael Schoenburg
    Last Edit: 26.03.2020
        
    This projects code loosely follows the PowerShell Practice and Style guide, as well as Microsofts PowerShell scripting performance considerations.
    Style guide: https://poshcode.gitbook.io/powershell-practice-and-style/
    Performance Considerations: https://docs.microsoft.com/en-us/powershell/scripting/dev-cross-plat/performance/script-authoring-considerations?view=powershell-7.1
.LINK
    https://github.com/MichaelSchoenburg/Get-AllAddresses
#>

#-----------------------------------------------------------------------------------------------------------------
#region RULES

# Rules/Conventions:

#     # This is a comment related to the following command

#     <# 
#         This is a comment related to the next couple commands (most times a caption for a section)
#     #>

# "This is a Text containing $( $variable(s) )"

# 'This is a text without variable(s)'

# Variables available in the entire script start with a capital letter
# $ThisIsAGlobalVariable

# Variables available only locally e. g. in a function start with a lower case letter
# $thisIsALocalVariable

#endregion RULES
#-----------------------------------------------------------------------------------------------------------------
#region INITIALIZATION

using namespace System.Management.Automation.Host

param (
    [Alias('Lang','L')]
    [Parameter(
        Mandatory = $false,
        Position = 1
    )]
    [ValidateSet('DE','EN')]
    [string]
    $global:Language
)

Add-Type -AssemblyName System.Windows.Forms
 
#endregion INITIALIZATION
#-----------------------------------------------------------------------------------------------------------------
#region DECLARATIONS

if (-not ($global:Language)) {
    switch ( $env:LANG ) {
        'de_DE.UTF-8' { $global:Language = 'DE' }
        Default { $global:Language = 'EN' }
    }
}

switch ( $global:Language ) {
    # German translations
    'DE' {
        $TConnecting = 'Verbinde zu'
        $TDataBeingPrepared = "Bitte haben Sie einen Moment Geduld. Ihre Daten werden gesammelt und aufbereitet."
        $TProcessing = 'Verarbeite'

        $TType = "Typ"

        $TSharedMailbox = "Freigegebenes Postfach"
        $TUserMailbox = "Regulaeres Postfach"
        $TMailUniversalDistributionGroup = "Verteiler"
        $TGroupMailbox = "Office 365-Gruppe"
        $TMailContact = "Kontakt"
        $TDiscoveryMailbox = "Dienstkonto"
        $TDynamicDistributionGroup = "Dynamischer Verteiler"
        $TEquipmentMailbox = "Equipment-Postfach"
        $TGuestMailUser = "Gast"
        $TLegacyMailbox = "Legacy-Postfach"
        $TLinkedMailbox = "Linked-Postfach"
        $TLinkedRoomMailbox = "Linked-Raum-Postfach"
        $TMailForestContact = "E-Mail-Stamm-Kontakt"
        $TMailNonUniversalGroup = "Nicht universelle Gruppe"
        $TMailUniversalSecurityGroup = "Universelle Sicherheits-Gruppe"
        $TMailUser = "E-Mail-Benutzer"
        $TPublicFolder = "Oeffentlicher Ordner"
        $TPublicFolderMailbox = "Oeffentlicher Ordner (E-Mail-Aktiviert)"
        $TRemoteEquipmentMailbox = "Remote-Equipment-Postfach"
        $TRemoteRoomMailbox = "Remote-Raum-Postfach"
        $TRemoteSharedMailbox = "Remote-Freigegebenes Postfach"
        $TRemoteTeamMailbox = "Remote-Team-Postfach"
        $TRemoteUserMailbox = "Remote-Benutzer-Postfach"
        $TRoomList = "Raumliste"
        $TRoomMailbox = "Raum-Postfach"
        $TSchedulingMailbox = "Planungspostfach"
        $TTeamMailbox = "Team-Postfach"

        $TDisplayName = "Anzeigename"
        $TPrimarySmtpAddress = "Primaere E-Mail-Adresse"

        $TQ1T = 'Anwendungsbereich'
        $TQ1 = 'Moechten Sie eine Tabelle aller E-Mail-Adressen aller Kunden oder alle Adressen eines spezifischen Kunden?'
        $TQ1A = 'Alle Kunden'
        $TQ1B = 'Spezifischer Kunde'
        
        $TH1T = 'Hinweis'
        $TH1 = 'Bitte geben Sie im folgenden die Zugangsdaten fuer einen delegierten Administrator an, der Zugriff auf alle Kunden hat.'

        $TOpeningOpenFolderDialogue = 'Oeffne browse file dialogue...'
        $TSelectAFolder = 'Waehlen Sie einen Ordner aus, in welchem Sie alle CSV-Dateien speichern wollen.'
        
        $TH2T = 'Hinweis'
        $TH2 = 'Bitte geben Sie im folgenden die Zugangsdaten fuer einen Adminstrator des online Exchanges an, von welchem Sie die E-Mail-Adressen auslesen moechten.'
        
        $TQ2T = 'Ausgabeart'
        $TQ2 = 'Moechten Sie die E-Mail-Adressen in eine CSV-Datei abspeichern oder direkt in einer interaktiven Tabelle anzeigen lassen, in welcher Sie Filtern, Sortieren und Spalten ein-/ausblenden koennen?'
        $TQ2A = 'Tabelle'
        $TQ2B = 'CSV-Datei'
    }

    # English translations
    'EN' {
        $TConnecting = 'Connecting to'
        $TDataBeingPrepared = "Please wait a moment. Your data is being prepared."
        $TProcessing = "Processing"

        $TType = "Type"

        $TSharedMailbox = "Shared Mailbox"
        $TUserMailbox = "Regular Mailbox"
        $TMailUniversalDistributionGroup = "Universal Distribution Group"
        $TGroupMailbox = "Office 365 Group"
        $TMailContact = "Contact"
        $TDiscoveryMailbox = "Service Mailbox"
        $TDynamicDistributionGroup = "Dynamic Distribution Group"
        $TEquipmentMailbox = "Equipment Mailbox"
        $TGuestMailUser = "Guest"
        $TLegacyMailbox = "Legacy Mailbox"
        $TLinkedMailbox = "Linked Mailbox"
        $TLinkedRoomMailbox = "Linked Room Mailbox"
        $TMailForestContact = "Forest Contact"
        $TMailNonUniversalGroup = "Non Universal Group"
        $TMailUniversalSecurityGroup = "Universal Security Group"
        $TMailUser = "Mail User"
        $TPublicFolder = "Public Folder"
        $TPublicFolderMailbox = "Public Folder (mail activated)"
        $TRemoteEquipmentMailbox = "Remote Equipment Mailbox"
        $TRemoteRoomMailbox = "Remote Room Mailbox"
        $TRemoteSharedMailbox = "Remote Shared Mailbox"
        $TRemoteTeamMailbox = "Remote Team Mailbox"
        $TRemoteUserMailbox = "Remote User Mailbox"
        $TRoomList = "Room List"
        $TRoomMailbox = "Room Mailbox"
        $TSchedulingMailbox = "Scheduling Mailbox"
        $TTeamMailbox = "Team Mailbox"

        $TDisplayName = "Displayname"
        $TPrimarySmtpAddress = "Primary SMTP Address"

        $TQ1T = 'Scope'
        $TQ1 = 'Do you need a table of all addresses from all customers or all addresses from one specific customer?'
        $TQ1A = 'All Customers'
        $TQ1B = 'Specific Customer'
        
        $TH1T = 'Info'
        $TH1 = 'Please provide credentials for an administrator account with delegated access to all customers after this.'
        
        $TOpeningOpenFolderDialogue = 'Opening browse folder dialog...'
        $TSelectAFolder = 'Select a folder to save all CSV files in'
        
        $TH2T = 'Info'
        $TH2 = 'In the next windows please provide credentials for an administrator inside the specific tenant which you want to scan.'
        
        $TQ2T = 'Output type'
        $TQ2 = 'Do you want to receive the output as an interactive table (you will be able to sort, filter and search) or in form of a CSV file?'
        $TQ2A = 'Table'
        $TQ2B = 'CSV file'
    }
}

#endregion DECLARATIONS
#-----------------------------------------------------------------------------------------------------------------
#region FUNCTIONS

function New-Menu {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Title, 

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Question, 

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ChoiceA, 

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ChoiceB
    )
    
    $a = [ChoiceDescription]::new("&$( $ChoiceA )", '')
    $b = [ChoiceDescription]::new("&$( $ChoiceB )", '')

    $options = [ChoiceDescription[]]($a, $b)

    $result = $host.ui.PromptForChoice($title, $question, $options, 0)

    return $result
}


function Write-ConsoleLog {
    <#
    .SYNOPSIS
        Logs an event to the console.
    
    .DESCRIPTION
        Writes text to the console with the current date in front of it.
    
    .PARAMETER Text
        Event/text to be outputted to the console.
    
    .EXAMPLE
        Write-ConsoleLog -Text 'Subscript XYZ called.'
        
        Long form
    .EXAMPLE
        Log 'Subscript XYZ called.
        
        Short form
    #>
    [Alias('Log')]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, 
        Position = 0)]
        [string]
        $Text
    )
    
    <# 
        Save current VerbosePreference
    #>
    $VerbosePreferenceBefore = $VerbosePreference
    
    <# 
        Enable verbose output
    #>
    # If I changed VerbosePreference at the star to the script
    # every function - even those from the Exchange Online module -
    # would add their own verbose output
    # This way I only have my own "verbose" output
    $VerbosePreference = 'Continue'

    <# 
        Write verbose output
    #>
    # (Verbose output doesn't interfere with function returns)
    # If this function was executed outside this script and thus without the
    # local variable $global:Language it would still work (but always choose english)
    if ($global:Language -eq 'DE') {
        Write-Verbose "$( Get-Date -Format "dd'.'MM'.'yyyy HH':'mm':'ss" ) - $( $text )"
    } else {
        Write-Verbose "$( Get-Date -Format "MM'/'dd'/'yyyy HH':'mm':'ss" ) - $( $text )"
    }
    
    <# 
        Restore current VerbosePreference
    #>
    $VerbosePreference = $VerbosePreferenceBefore
}

function Get-AllTenantAddresses {
    [CmdletBinding(DefaultParameterSetName = 'UseActiveSession')]
    param (
        [Parameter(
            ParameterSetName = 'UseCustomTenant',
            Mandatory = $true, 
            Position = 1, 
            HelpMessage = 'The onmicrosoft-domain identifying the custom tenant to connect to.'
        )]
        [ValidateScript({$_ -like '*.*'})]
        [string]
        $TenantId,

        [Parameter(
            ParameterSetName = 'UseActiveSession',
            Position = 1,
            HelpMessage = 'Use an already connected session.'
        )]
        [switch]
        $UseActiveSession = $false
    )

    process {
        # Declarations
        $finalAddresses = @()
        $aliasCounter = @()

        if (-not ($UseActiveSession)) {
            try {
                <# 
                    Connect to EXO
                #>
                Log "$( $TConnecting ) $( $TenantId )..."
                # Suppressing output by using null
                $null = Connect-ExchangeOnline -DelegatedOrganization $TenantId -ShowBanner:$false            
            }
            catch {
                if ($_.CategoryInfo.Activity -eq 'New-ExoPSSession') {
                    Log "Unable to connect to $( $domain )."
                    # Return error code 1
                    return 1
                } else {
                    # Return error code 2
                    return 2
                }
            }
        }

        <# 
            Get all entities present on the EXO
        #>
        try {
            Log $TDataBeingPrepared 
            $Mbxs = Get-Recipient

            # Classify each entity
            for ($i = 0; $i -lt $Mbxs.Count; $i++) {
                Log "[$( $i + 1 )/$( $Mbxs.Count )] $( $TProcessing ) $( $Mbxs[$i].PrimarySmtpAddress )"

                $MbxAdrFiltered = New-Object PSObject
                $MbxAdrUnfiltered = Get-Recipient $Mbxs[$i].Identity | Select-Object -ExpandProperty EmailAddresses
                $ACount = 0
                $EntityType = $Mbxs[$i].RecipientTypeDetails

                # Define which type of entity it is
                switch ($EntityType) {
                    "SharedMailbox" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TSharedMailbox; Break }
                    "UserMailbox" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TUserMailbox; Break }
                    "MailUniversalDistributionGroup" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TMailUniversalDistributionGroup; Break }
                    "GroupMailbox" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TGroupMailbox; Break }
                    "MailContact" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TMailContact; Break }
                    "DiscoveryMailbox" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TDiscoveryMailbox; Break }
                    "DynamicDistributionGroup" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TDynamicDistributionGroup; Break }
                    "EquipmentMailbox" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TEquipmentMailbox; Break }
                    "GuestMailUser" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TGuestMailUser; Break }
                    "LegacyMailbox" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TLegacyMailbox ; Break }
                    "LinkedMailbox" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TLinkedMailbox; Break }
                    "LinkedRoomMailbox" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TLinkedRoomMailbox; Break }
                    "MailForestContact" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TMailForestContact; Break }
                    "MailNonUniversalGroup" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TMailNonUniversalGroup; Break }
                    "MailUniversalSecurityGroup" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TMailUniversalSecurityGroup; Break }
                    "MailUser" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TMailUser; Break }
                    "PublicFolder" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TPublicFolder; Break }
                    "PublicFolderMailbox" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TPublicFolderMailbox; Break }
                    "RemoteEquipmentMailbox" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TRemoteEquipmentMailbox; Break }
                    "RemoteRoomMailbox" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TRemoteRoomMailbox; Break }
                    "RemoteSharedMailbox" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TRemoteSharedMailbox; Break }
                    "RemoteTeamMailbox" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TRemoteTeamMailbox; Break }
                    "RemoteUserMailbox" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TRemoteUserMailbox; Break }
                    "RoomList" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TRoomList; Break }
                    "RoomMailbox" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TRoomMailbox; Break }
                    "SchedulingMailbox" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TSchedulingMailbox; Break }
                    "TeamMailbox" { $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TType -Value $TTeamMailbox; Break }
                }
            
                # Add displayname
                $MbxAdrFiltered | Add-Member -type NoteProperty -Name $TDisplayName -Value $Mbxs[$i].DisplayName
                
                # Search for primary SMTP address
                $PrimarySmtpAddr = ($MbxAdrUnfiltered.where({$_ -clike '*SMTP*'}))[0]
                $MbxAdrFiltered | Add-Member -type NoteProperty -Name 'Primaere E-Mail-Adresse' -Value $PrimarySmtpAddr.Substring(5)
            
                # Search for aliases
                $Aliases = $MbxAdrUnfiltered.where({$_ -clike '*smtp*'})
                ForEach ($e in $Aliases) {
                    $ACount++
                    $MbxAdrFiltered | Add-Member -type NoteProperty -Name "Alias$( $ACount )" -Value $e.Substring(5)
                }
            
                $finalAddresses += $MbxAdrFiltered
                $aliasCounter += $ACount
            }
            
            # Prepare output
            $properties = $TType, $TDisplayName, $TPrimarySmtpAddress
            # Get the highest count of aliases per entity there is
            $aliasCounter = $aliasCounter | Sort-Object -Descending
            For ($i = 1; $i -le $aliasCounter[0]; $i++) {
                $properties += "Alias$($i)"
            }
            # Sort the output
            $finalAddresses = $finalAddresses | Sort-Object -Property $TType, $TDisplayName

            if (-not ($UseActiveSession)) {
                # Close EXO session(s)
                Log "Closing EXO session..."
                foreach($s in (Get-PSSession).where({$_.ComputerName -eq 'outlook.office365.com'})){
                    Remove-PSSession -Id $s.Id
                }
            }

            # Finally return output
            return $finalAddresses
        }
        catch {
            Log 'Error processing entries:'
            Log $_.Exception.Message
            # Return error Code 3
            return 3
        }
    }
}

#endregion FUNCTIONS
#-----------------------------------------------------------------------------------------------------------------
#region EXECUTION

# Decide whether to process all customers or just a specific one
$ResultCustomer = New-Menu -Title $TQ1T -Question $TQ1 -ChoiceA $TQ1A -ChoiceB $TQ1B
switch ($ResultCustomer) {
    0 {
        #region ALLCUSTOMERS

        <# 
            Set the location where to save the output
        #>
        Log $TOpeningOpenFolderDialogue
        # Suppress output by using null
        $null = [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        $DialogueOpenFolder = New-Object System.Windows.Forms.FolderBrowserDialog
        $DialogueOpenFolder.Description = $TSelectAFolder
        $DialogueOpenFolder.RootFolder = "MyComputer"
        # initial folder
        $DialogueOpenFolder.SelectedPath = "$( $HOME )\Desktop"
        $DialogueOpenFolder.ShowDialog()
        $Folder = $DialogueOpenFolder.SelectedPath

        # Get all customers tenant IDs
        # Console text messages are often overseen when the connection request form appears,
        # so I decided to show a Windows Message Form that can not be overseen.
        # Suppress output by using null
        $null = [System.Windows.Forms.MessageBox]::Show(
            # Text
            $TH1,
            # Title
            $TH1T,
            # Button
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            # idk
            [System.Windows.Forms.MessageBoxIcon]::Asterisk
        )
        Connect-AzureAD
        
        # Get all customers delegated access is set up for
        $Domains = (Get-AzureADContract).DefaultDomainName

        # Get all addresses for each customer
        ForEach( $domain in $Domains ){
            $AllAddresses = Get-AllTenantAddresses -TenantId $domain
            
            if ($AllAddresses.GetType() -eq [Int]) {
                # If something went wront
                Log "Skipping $( $domain )."
            } if ( $AllAddresses.GetType() -eq [Object[]] ) {
                <# 
                    Save output to CSV file
                #>
                $Path = "$( $Folder )\$( $domain ).csv"
                $properties = $TType, $TDisplayName, $TPrimarySmtpAddress
                # Export-Csv -InputObject doesn't take Arrays so it has to be solved via pipe
                $AllAddresses | Export-Csv -NoTypeInformation -Delimiter ';' -Path $Path -Encoding UTF8
            }
        }

        #endregion ALLCUSTOMERS
    }
    1 {
        #region SPECIFICCUSTOMER
        
        [System.Windows.Forms.MessageBox]::Show($TH2, $TH2T, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Asterisk) | Out-Null
        Connect-ExchangeOnline -ShowBanner:$false
        
        $ResultOutput = New-Menu -Title $TQ2T -ChoiceA $TQ2A -ChoiceB $TQ2B -Question $TQ2
        
        # Daten erfassen
        $finalAddresses = Get-AllTenantAddresses -UseActiveSession

        # Ausgabe auf Basis der zuvor gewaehlten Ausgabemethode durchfuehren
        switch( $ResultOutput ){
            0 {
                # GUI
                $finalAddresses | Select-Object -Property $Properties | Out-GridView
            }
            1 {
                # CSV
                $Dialog = New-Object System.Windows.Forms.SaveFileDialog
                $Dialog.Filter = "CSV-Dateien (*.csv)|*.csv|All Files (*.*)|*.*"
                $Dialog.ShowDialog()
                $Dialog.SupportMultiDottedExtensions = $true;
                $Dialog.InitialDirectory = "$( $HOME )\Desktop"
                $Dialog.CheckFileExists = $true
                $finalAddresses | Select-Object -Property $Properties | Export-Csv -NoTypeInformation -Delimiter ";" -Path $Dialog.FileName
            }
        }

        #endregion SPECIFICCUSTOMER
    }
}

#endregion EXECUTION
#-----------------------------------------------------------------------------------------------------------------
#region EXECUTION
