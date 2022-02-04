Add-Type -Path "C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.3.1.0\lib\netstandard2.0\MimeKit.dll"
Add-Type -Path "C:\Program Files\PackageManagement\NuGet\Packages\MailKit.3.1.0\lib\netstandard2.0\MailKit.dll"

function Send-MailKitMessage{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [string]$From,
        [Parameter(Mandatory)]
        $To,
        [Parameter()]
        $CC,
        [Parameter()]
        $BCC,
        [Parameter()]
        [string]$Subject="",
        [Parameter()]
        [string]$Body="",
        [Parameter()]
        $Attachments,
        [Parameter(Mandatory)]
        [string]$SMTPServer,
        [Parameter()]
        [int32]$Port=25,
        [Parameter()]
        [switch]$BodyAsHtml,
        [Parameter()]
        $Credential
    )

    $SMTP=New-Object MailKit.Net.Smtp.SmtpClient
    $Message=New-Object MimeKit.MimeMessage
    $Builder=New-Object MimeKit.BodyBuilder

    $Message.From.Add($From)

    foreach($Person in $To){
        $Message.To.Add($Person)
    }

    if($CC){
        foreach($Person in $CC){
            $Message.Cc.Add($Person)
        }
    }

    if($BCC){
        foreach($Person in $BCC){
            $Message.Bcc.Add($Person)
        }
    }

    $Message.Subject=$Subject

    if($BodyAsHtml){
        $Builder.HtmlBody=$Body
    }else{
        $Builder.TextBody=$Body
    }

    if($Attachments){
        foreach($Attachment in $Attachments){
            $Builder.Attachments.Add($Attachment)
        }
    }

    $Message.Body=$Builder.ToMessageBody()

    $SMTP.Connect($SMTPServer,$Port,$false)

    if($Credential){
        $SMTP.Authenticate($Credential.username,$Credential.getNetworkCredential().password)
    }

    $SMTP.Send($Message)

    $SMTP.Disconnect($true)
    $SMTP.Dispose()
}
