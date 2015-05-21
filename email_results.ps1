$mailprefs = @{
            'To'           = 'jsommerville@bayportstatebank.com';
            'Subject'      = '3rd Party Update Report';
            'Body'         = 'No Report Found!';
            'From'         = 'ninite@bayportstatebank.com';
            'SmtpServer'   = 'bayportstatebank-com.mail.protection.outlook.com'
}

if(Test-Path .\UpdateReport.html) {
    $body = Get-Content .\UpdateReport.html | Out-String
    $mailprefs.Set_Item('Body',$body)
    $mailprefs.Add('BodyAsHtml',$true)
}

Send-MailMessage @mailprefs