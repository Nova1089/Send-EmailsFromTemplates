<#
This script sends out a set of emails using templates created in Outlook.
To create an email template, follow the steps in this article:
https://support.microsoft.com/en-us/office/create-an-email-message-template-43ec7142-4dd0-4351-8727-bd0977b6b2d1
Place all templates that you want to send into the same folder.
#>

# functions
function Show-Introduction
{
    Write-Host ("This script sends out a set of emails using templates created in Outlook.`n" +
        "To create an email template, follow the steps in this article:`n" +
        "https://support.microsoft.com/en-us/office/create-an-email-message-template-43ec7142-4dd0-4351-8727-bd0977b6b2d1`n" +
        "Place all templates that you want to send into the same folder.`n") -ForegroundColor DarkCyan
}

function Test-OutlookAlreadyOpen
{
    return ($null -ne (Get-Process Outlook -ErrorAction SilentlyContinue))
}

function Test-ValidContext($outlookAlreadyOpen)
{
    if ($outlookAlreadyOpen)
    {
        if (Test-SessionIsAdmin)
        {
            Throw ("Outlook is already open (most likely without admin priveleges) and Powershell is running with admin priveleges.`n" +
                "Powershell cannot use the Outlook session with admin priveleges.`n" +
                "Either close Outlook, or open Powershell WITHOUT admin privileges, and try again.")
        }
    }
}

function Test-SessionIsAdmin
{
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentSessionIsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    return $currentSessionIsAdmin
}

function Get-TemplatesFolder
{
    do
    {
        $templatesFolderPath = Read-Host "Enter path to folder containing templates (i.e. C:\Templates)"
        $templatesFolderPath = $templatesFolderPath.Trim('"') # trim quotes if they were included
        $folder = Get-Item -Path $templatesFolderPath -ErrorAction SilentlyContinue
        
        if ($null -eq $folder)
        {
            $folderExists = $false
            Write-Warning "Folder not found. Please try again."
        }
        else
        {
            $folderExists = $true
        }
    }
    while (-not($folderExists))

    return $templatesFolderPath
}

function Get-TemplatePaths($folderPath)
{
    do
    {
        $templatePaths = TryGet-TemplatePaths $folderPath
        
        if ($null -eq $templatePaths)
        {
            Write-Warning "No email templates were found in this folder. Please add .oft template files and try again."
            Read-Host "Press Enter to continue"
        }
    }
    while ($null -eq $templatePaths)

    return $templatePaths
}

function TryGet-TemplatePaths($folderPath)
{
    $allFiles = Get-ChildItem -Path $folderPath -Recurse
    $templatePaths = New-Object System.Collections.Generic.List[string]
    $totalTemplates = 0

    foreach ($file in $allFiles)
    {
        if ($file.Extension -eq ".oft")
        {
            $totalTemplates++
            $templatePaths.Add($file.FullName)
        }
    }
    
    if ($totalTemplates -ne 0)
    {
        Write-Host "Found $totalTemplates email templates." -ForegroundColor DarkCyan
    }

    return $templatePaths
}

function Open-OutlookSession
{
    return New-Object -ComObject Outlook.Application
}

function Send-AllEmails($outlookSession, $templatePaths)
{
    Write-Host "Sending emails..." -ForegroundColor DarkCyan
    $totalEmailsSent = 0

    foreach ($templatePath in $templatePaths)
    {
        $templateName = Split-Path -Path $templatePath -Leaf

        try
        {
            Send-Email -outlookSession $outlookSession -templatePath $templatePath
        }
        catch
        {
            Write-Warning "$templateName failed to send. It may be missing something crucial (i.e the send address)."
            continue
        }
        
        $totalEmailsSent++        
        Write-Host "Email template was sent: $templateName" -ForegroundColor Green        
    }

    Write-Host "Total emails sent: $totalEmailsSent`n" -ForegroundColor Green
}

function Send-Email($outlookSession, $templatePath)
{
    $message = $outlookSession.CreateItemFromTemplate($templatePath)
    $message.Send()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($message) | Out-Null
}

function Close-OutlookSession($outlookSession, $outlookAlreadyOpen)
{
    if (-not($outlookAlreadyOpen))
    {
        $outlookSession.Quit()
    }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlookSession) | Out-Null
}

# main
Show-Introduction
$outlookAlreadyOpen = Test-OutlookAlreadyOpen
Test-ValidContext $outlookAlreadyOpen
$templatesFolderPath = Get-TemplatesFolder
$templatePaths = Get-TemplatePaths $templatesFolderPath
Read-Host "Press Enter to send the emails"
$outlookSession = Open-OutlookSession
Send-AllEmails -outlookSession $outlookSession -templatePaths $templatePaths
Close-OutlookSession -outlookSession $outlookSession -outlookAlreadyOpen $outlookAlreadyOpen
Read-Host "Press Enter to exit"