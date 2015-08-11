#########################
# Title    - ADComputerObjectCheck
# Author   - jwingram
# Purpose  - Checks the Default Computers OU for objects that do not belong.
# Date     - 12/18/12 updated 8/11/15
#########################
#Checks if Quest ActiveRoles is already loaded.  If not, load the Snapin.  If the software isn't installed, prompt.
$Quest = Get-PSSnapin Quest.ActiveRoles.ADManagement -ErrorAction silentlycontinue
if (!$Quest)
  {
  Write-host "Loading Quest.ActiveRoles.ADManagement Snapin"
  Add-PSSnapin Quest.ActiveRoles.ADManagement
  if (!$?) 
    {"Need to install AD Snapin from http://www.quest.com/powershell";exit}
  }
    
#########################
# SCRIPT START
#########################
$sname=$myInvocation.MyCommand.Definition
$DomainsToCheck= "domain.name"

$emailFrom = "from@email.com"
$emailTo = "to@email.com"
$subject = "Computer Objects in default Container"
$smtpServer = "smtp.office365.com"
$passwd = ConvertTo-SecureString “PASSWORD” -AsPlainText -Force
$credentials = New-Object System.Management.Automation.PSCredential (“from@email.com”, $passwd)

ForEach ($Domain in $DomainsToCheck)
  {
  $garbage=connect-QADService $domain
  $results=get-qadcomputer -searchroot "$domain/computers" | Select-object name, operatingsystem

  if ($results -ne $null)
    {
    $body = "The following computer objects appear to be in the default 'Computers' container in the $domain domain. Please move these to the correct OU.<br><br>"
    $body+="<font color=red>"
    Foreach ($result in $results)
      {
      $body+=$result.Name + "<br>"
      }
    $body+="<br><font color=black><p style=`"font-size:x-small;`">*****************************<br>This is an automated script running at $sname on $env:computername as $env:username"

    send-mailmessage -Credential $credentials -UseSsl -Port "587" -to $emailTo -from $emailFrom -subject $subject -smtpserver $smtpServer -body $body -bodyashtml
    }
  }