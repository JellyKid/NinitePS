$mailprefs = @{
            'To'           = 'jsommerville@bayportstatebank.com';
            'Subject'      = 'Ninite Reports';
            'Body'         = 'No Report Found!';
            'From'         = 'ninite@bayportstatebank.com';
            'SmtpServer'   = 'bayportstatebank-com.mail.protection.outlook.com'
}

Push-Location $PsScriptRoot

if(Test-Path .\Report.html) {
    $body = Get-Content .\Report.html | Out-String
    $mailprefs.Set_Item('Body',$body)
    $mailprefs.Add('BodyAsHtml',$true)
}

Send-MailMessage @mailprefs

Pop-Location