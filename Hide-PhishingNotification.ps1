function Hide-PhishingNotification {
<#
.SYNOPSIS 
Use Outlook client to create a server rule that will "delete" phishing notification emails.
 
.DESCRIPTION 
This script will use a COM object to control the outlook client to create a rule that checks the body and subject
of all emails for strings such as "phishing" and "hacked" and moves it to the deleted items folder if matched. 
There is also an option to specify a specific from email address, such as "itsupport" or "security". 

Furthermore, this script provides the ability to delete the rule to clean up. 

Note: There is no way to permantently delete an email using rules, so the phishing notification will be moved to the deleted folder instead. 
TODO: add more fine grained rules, such as when subject includes "notification" AND body contains "clicked". Would require multiple rules. 
TODO: add capability to check and permanently delete phishing notification 
TODO: Port to js or vbs
VBS reference: https://msdn.microsoft.com/en-us/library/bb206765(v=office.12).aspx

.PARAMETER RuleName
Name of rule

.PARAMETER FromEmail
Only apply rule to this email address. Useful when all phishing notifications come from a certain email, like itsupport.

.PARAMETER BodyOrSubjectWords
Words to filter on. It will be triggered if the body or subject contains any of those words. There is no AND option in Outlook. 

.EXAMPLE 
PS > Hide-PhishingNotification
Create rule named "{Firstname}'s Rule" that will move emails to deleted folder if body or subject matches words "phish","hack","malware","security incident"

#>
  Param([string]$RuleName = (Get-Culture).TextInfo.ToTitleCase(($env:UserName -split " ")[0]) + "'s Rule", #ex: John's Rule
        [string]$FromEmail, 
        [string[]]$BodyOrSubjectWords = @("phish","hack","malware","security incident","scam"), 
        [switch]$Cleanup) #delete rule
    #Code template from: 
    #http://dandarache.wordpress.com/2011/07/25/using-powershell-to-create-rules-in-outlook/
    function Add-OutLookRule
    {
        param([string]$RuleName,
              [string]$FromEmail,
              [string]$ForwardEmail,
              [string]$RedirectFolder,
              [switch]$Delete,
              [switch]$RemoveRule,
              [string[]]$SubjectWords,
              [string[]]$BodyWords,
              [string[]]$BodyOrSubjectWords)
        Add-Type -AssemblyName microsoft.office.interop.outlook 
        $olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]
        $olRuleType = "Microsoft.Office.Interop.Outlook.OlRuleType" -as [type]
        $outlook = New-Object -ComObject outlook.application
        $namespace  = $Outlook.GetNameSpace("MAPI")
        $inbox = $namespace.getDefaultFolder($olFolders::olFolderInbox)
        $rules = $outlook.session.DefaultStore.GetRules()
        if($RemoveRule) {
            $rules.Remove($RuleName)
            $rules.Save()
            return
        }
        $rule = $rules.Create($RuleName,$olRuleType::OlRuleReceive)
        if ($SubjectWords) {
            $SubjectCondition = $rule.Conditions.Subject
            $SubjectCondition.Enabled = $true
            $SubjectCondition.Text = $SubjectWords
        }
        if ($BodyOrSubjectWords) {
            $BodyOrSubjectCondition = $rule.Conditions.BodyOrSubject
            $BodyOrSubjectCondition.Enabled = $true
            $BodyOrSubjectCondition.Text = $BodyOrSubjectWords
        }
        if ($BodyWords) {
            $BodyCondition = $rule.Conditions.Body
            $BodyCondition.Enabled = $true
            $BodyCondition.Text = $BodyWords
        }
       if ($RedirectFolder) {
           $d = [System.__ComObject].InvokeMember(
                "EntryID",
                [System.Reflection.BindingFlags]::GetProperty,
                $null,
                $inbox.Folders.Item($RedirectFolder),
                $null)#>
                $MoveTarget = $namespace.getFolderFromID($d)
       } elseif ($delete){
            $MoveTarget = $namespace.getDefaultFolder(
            $olFolders::olFolderDeletedItems)
       }
       if ($MoveTarget) {
            $MoveRuleAction = $rule.Actions.MoveToFolder
            [Microsoft.Office.Interop.Outlook._MoveOrCopyRuleAction].InvokeMember(
                "Folder",
                [System.Reflection.BindingFlags]::SetProperty,
                $null,
                $MoveRuleAction,
                $MoveTarget)
            $MoveRuleAction.Enabled = $true
        }
        if ($FromEmail) {
            $FromCondition = $rule.Conditions.From
            $FromCondition.Enabled = $true
            $FromCondition.Recipients.Add($FromEmail)
            $fromCondition.Recipients.ResolveAll()
        }
        if ($ForwardEmail) {
            $ForwardRuleAction = $rule.Actions.Forward
            $ForwardRuleAction.Recipients.Add($ForwardEmail)
            $ForwardRuleAction.Recipients.ResolveAll()
            $ForwardRuleAction.Enabled = $true
        }
        <# Delete action simply moves it to delete items. Might as well move it to delete items instead to be less suspicious
        if ($delete) {
            $deleteRule = $rule.Actions.Delete
            $deleteRule.Enabled = $true
        }#>
        $rules.Save()
    }

    if($Cleanup) {
        Add-OutLookRule -RemoveRule -RuleName $RuleName
    } else {
        Add-OutLookRule -RuleName $RuleName -BodyOrSubjectWords $BodyOrSubjectWords -FromEmail $FromEmail -Delete 
    }
}

