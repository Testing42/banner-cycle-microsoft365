#Script rotates email banner notification colors to help ensure they are better noticed.

 

#importing exchange

Import-Module ExchangeOnlineManagement

 

#setting variables

$cert = Get-AutomationCertificate -Name 'WindowsAutomationApp'

$AppID = Get-AutomationVariable -Name 'App-AutomationApplication-AppId'

 

#connecting to exchangeonline.

Connect-ExchangeOnline -Certificate $cert -AppID $AppID -Organization "tenantid"

 

#Set transport rule for external banner

$Gold = Get-TransportRule "Visual Cue - External to Organization" | Select State

$Blue = Get-TransportRule "Visual Cue - External to Organization 2" | Select State

$Purple = Get-TransportRule "Visual Cue - External to Organization 3" | Select State

 

if ($Gold.State -eq 'Enabled') {

    Disable-TransportRule -Identity "Visual Cue - External to Organization" -Confirm:$false

    Enable-TransportRule -Identity "Visual Cue - External to Organization 2" -Confirm:$false

    }

elseif ($Blue.State -eq 'Enabled') {

    Disable-TransportRule -Identity "Visual Cue - External to Organization 2" -Confirm:$false

    Enable-TransportRule -Identity "Visual Cue - External to Organization 3" -Confirm:$false

   }

elseif ($Purple.State -eq 'Enabled') {

    Disable-TransportRule -Identity "Visual Cue - External to Organization 3" -Confirm:$false

    Enable-TransportRule -Identity "Visual Cue - External to Organization" -Confirm:$false

    }

 

#Set transport rule for SPF check banner

$Brick = Get-TransportRule "Visual Cue - No SPF Validation" | Select State

$Red = Get-TransportRule "Visual Cue - No SPF Validation 2" | Select State

$Orange = Get-TransportRule "Visual Cue - No SPF Validation 3" | Select State

 

if ($Brick.State -eq 'Enabled') {

    Disable-TransportRule -Identity "Visual Cue - No SPF Validation" -Confirm:$false

    Enable-TransportRule -Identity "Visual Cue - No SPF Validation 2" -Confirm:$false

    }

elseif ($Red.State -eq 'Enabled') {

    Disable-TransportRule -Identity "Visual Cue - No SPF Validation 2" -Confirm:$false

    Enable-TransportRule -Identity "Visual Cue - No SPF Validation 3" -Confirm:$false

   }

elseif ($Orange.State -eq 'Enabled') {

    Disable-TransportRule -Identity "Visual Cue - No SPF Validation 3" -Confirm:$false

    Enable-TransportRule -Identity "Visual Cue - No SPF Validation" -Confirm:$false

}

 

Disconnect-ExchangeOnline -Confirm:$false
