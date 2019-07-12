Try {
  Write-Host "Importing Modules"
  Import-Module ActiveDirectory
  Import-Module $PSScriptRoot\lib\WPFBot3000\0.9.25\WPFBot3000.psm1 -Verbose
  Add-Type -Path (Join-Path -Path (Split-Path $script:MyInvocation.MyCommand.Path) -ChildPath 'lib\CubicOrange.Windows.Forms.ActiveDirectory.dll')
}catch{
  "Something bad happened"
  break
}

#region  Begin function definitions 

  Function confirm {
    param(
        [string]$message,
        [string]$question
    )

    $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes'))
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))

    $decision = $Host.UI.PromptForChoice($message, $question, $choices, 1)
    if ($decision -eq 0) {
        return $True
    } else {
        return $False
    }

  }

  Function Get-ADPath {

    $DialogPicker = New-Object CubicOrange.Windows.Forms.ActiveDirectory.DirectoryObjectPickerDialog

    $DialogPicker.AllowedLocations = [CubicOrange.Windows.Forms.ActiveDirectory.Locations]::All
    $DialogPicker.AllowedObjectTypes = [CubicOrange.Windows.Forms.ActiveDirectory.ObjectTypes]::Groups,[CubicOrange.Windows.Forms.ActiveDirectory.ObjectTypes]::Users,[CubicOrange.Windows.Forms.ActiveDirectory.ObjectTypes]::Computers
    $DialogPicker.DefaultLocations = [CubicOrange.Windows.Forms.ActiveDirectory.Locations]::JoinedDomain
    $DialogPicker.DefaultObjectTypes = [CubicOrange.Windows.Forms.ActiveDirectory.ObjectTypes]::Users
    $DialogPicker.ShowAdvancedView = $false
    $DialogPicker.MultiSelect = $true
    $DialogPicker.SkipDomainControllerCheck = $true
    $DialogPicker.Providers = [CubicOrange.Windows.Forms.ActiveDirectory.ADsPathsProviders]::Default

    $DialogPicker.AttributesToFetch.Add('samAccountName')
    $DialogPicker.AttributesToFetch.Add('title')
    $DialogPicker.AttributesToFetch.Add('department')
    $DialogPicker.AttributesToFetch.Add('distinguishedName')


    [void]$DialogPicker.ShowDialog()

    If ($DialogPicker.SelectedObject.FetchedAttributes) {
      return $DialogPicker.Selectedobject.FetchedAttributes[3]
    }
  }

  #  Generates window and proccesses input
  Function main {

    $Dialog = Dialog {
    TextBox FirstName
    TextBox LastName
    TextBox Ticket_Number
    TextBox Title
    TextBox Department
    
    #  Section for Manager search controls 
    TextBox Manager    
    Button Search `
    -Property @{ `
      MinHeight = 26; MinWidth = 50; 
      } `
    -Action {
      $txt=$this.Window.GetControlByName('Manager')
      $found = Get-ADPath
      $txt.Text=$found
    }
    # end manager controls 

    Separator -Property @{ MinHeight = 15}
    ComboBox Org_Unit -Contents (Get-ADOrganizationalUnit -Filter 'Name -like "*"').Name
    TextBox Password "Password123"
    Separator -Property @{ MinHeight = 15}
    Checkbox Enabled
    Checkbox ChangePasswordAtLogon
    Separator
    Combobox O365_Licence -contents 'O365 E3','O365 Business Premium','Business Essentials','No Licence' -initialvalue 'No Licence'
    Separator

    } -property @{ Title = "New Wheelhouse AD User"; MinHeight = 144; MinWidth = 533; }

    if ($Dialog) {
      #  Hash table to pass to New-Aduser
      [PSCustomObject]$user = [ordered]@{
        Name                  = "$($Dialog.firstname) $($Dialog.lastname)"
        DisplayName           = "$($Dialog.firstname) $($Dialog.lastname)"
        GivenName             = $($Dialog.firstname)
        Surname               = $($Dialog.lastname)
        Manager               = $($Dialog.Manager)
        SamAccountName        = $Dialog.firstname.ToLower()[0] + $Dialog.LastName.ToLower()[0] + $Dialog.Ticket_Number #(-join (0..9 | get-random -Count 3))
        Path                  = Get-ADOrganizationalUnit -Filter ('Name -like "{0}"' -f ($dialog.Org_Unit))
        AccountPassword       = (ConvertTo-SecureString $Password -AsPlainText -Force)
        UserPrincipalName     = "$($Dialog.firstname).$($Dialog.lastname)@wheelhouse.solutions"
        EmailAddress          = "$($Dialog.firstname).$($Dialog.lastname)@wheelhouse.solutions"
        ChangePasswordAtLogon = $Dialog.ChangePasswordAtLogon
        Enabled               = $Dialog.Enabled
        Title                 = $Dialog.Title
        Company               = "Wheelhouse Solutions"
        Department            = $Dialog.Department
      }


      Write-host "`n:.New User details.:`n" -ForegroundColor Magenta
      $user
      Write-host "`n365 Licence type: $($Dialog.O365_Licence)`n" -ForegroundColor Green

      $cli = "New-ADUser @user"
      If (confirm -message "`nNew-ADUser" -question "Would you like to continue with this action?") {
        iex $cli
      }else{
        Write-Host "`nUser has not been created`n"
      }
      #  Start-Sleep -Seconds 5 #  breathing room
      #  Add-ADGroupMember -Identity $($user.SamAccountName) -Member "$($Dialog.O365_Licence)" -Confirm


      #  send the email
      Function SendEmail {
      $output = $user | ConvertTo-Html -As List | Out-String
      $Outlook = New-Object -ComObject Outlook.Application
      $Mail = $Outlook.CreateItem(0)
      $Mail.To = "tac@itservices.team"
      $Mail.Subject = "Service Ticket #$($Dialog.Ticket_Number)"
      $Mail.HTMLBody = $output
      $Mail.Send()
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
      }

      <#
        try {
        If ($u = Get-ADUser $user.SamAccountName -ErrorAction Stop) {return $u}

        }
        catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] { 

        Write-Warning "`nUser was not created"

        }
      #>
    }

  }
#endregion Define functions

#  start Form
main


$msg = "Do you want to create another user? [Y/N]"

do {
    $response = Read-Host -Prompt $msg
    if ($response -eq 'y') {
        main
    }
} until ($response -eq 'n')