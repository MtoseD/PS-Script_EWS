<# run on Windows Server #>


#global variables
$email    = "YOUR EMAIL"
$username = "YOUR MAILNAME" 
$password = "YOUR PASSWORK"
$domain   = "YOUR DOMAIN"
$USER_DEFINED_FOLDER_IN_MAILBOX = "inbox" #Mailbox you want
$date = Get-Date -Format "yyyy-MM-dd"
$dateCsv = Get-Date -Format "MM.yyyy"

#handle csv
function handleCsv
{
    #read csv from mail
    $downloadDirectory = 'YOUR PATH'
    $File = new-object System.IO.FileStream(($downloadDirectory + "\" + $attachment.Name.ToString()), [System.IO.FileMode]::Create)
	$File.Write($attachment.Content, 0,$attachment.Content.Length)
	$File.Close()         
    $csvFile = $downloadDirectory + "\” + $attachment.Name.ToString()
    $csv = Import-Csv -Delimiter ";" -Path "$csvFile" | select -Property "YOUR PROPERTY 1","YOUR PROPERTY 2","YOUR PROPERTY 3"
    $csvNr = Import-Csv -Delimiter ";" -Path "$csvFile" | select -Property "YOUR PROPERTY 1","YOUR PROPERTY 2"
    #filter csv
    foreach($nr in $csvNr)
    {
        if(([double]$nr.'YOUR PROPERTY 1' -gt 1000)-and($nr.'YOUR PROPERTY 2' -eq $dateCsv)) #you can set your own number at "1000"
        {
            #mail-variables
            $csvMail = $csv | Where-Object -Property 'YOUR PROPERTY 1' -EQ $nr.'YOUR PROPERTY 1'
            $name = $csvMail."YOUR PROPERTY 3" #if name is in csv-file
            #mail-message
            $message = 
            "
            Hello there <br><br>
            This is html-based
            "
            #write & send mail
            $mail = New-Object -TypeName Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $s
            $mail.Subject = 'YOUR SUBJECT'
            $mail.Body = $message
            $mail.ToRecipients.Add("$name@domain.ch") | Out-Null #or set static mail
            $mail.SendAndSaveCopy()
            echo "mail has been sent"
        }
        else{
            echo "doesn't match"
        }
    }
}

#check & set ExecutionPolicy
function Check-ExecutionPolicy{
    
    $check = Get-ExecutionPolicy

    if(!($check -eq 'RemoteSigned' -or 'Unrestricted'))
    {
        Set-ExecutionPolicy RemoteSigned
    }
}

#call function
Check-ExecutionPolicy

#Web Services
$EXCHANGE_WEB_SERVICE_DLL = "YOUR-PATH \Web Services\2.2\Microsoft.Exchange.WebServices.dll"

#load the assembly
[void] [Reflection.Assembly]::LoadFile($EXCHANGE_WEB_SERVICE_DLL)

#set reference to exchange
$s = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService

#use first option if you want to impersonate, otherwise, grab your own credentials
$s.Credentials = New-Object Net.NetworkCredential($username, $password, $domain)

#discover the url from your email address
$s.AutodiscoverUrl($email)

#handle inbox
$inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($s,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
$MailboxRootid = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $email) #selection and creation of new root
$MailboxRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($s,$MailboxRootid)

#filter inbox
$searchfilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject,"YOUR SUBJECT")     
$itemView = new-object Microsoft.Exchange.WebServices.Data.ItemView(999)
$searchResults = $s.FindItems($MailboxRoot.Id, $searchfilter, $itemView)

#handle result
foreach($result in $searchResults)
{   
    #check date
    if($result.DateTimeSent -ge $date)
    {
        #check attachment
        if($result.HasAttachments -eq $true)
        {
            $result.Load()

            #check file format
            foreach($attachment in $result.Attachments)
            {
                #checks if it is a .csv-file
                if($attachment.Name.Contains(".csv"))
                {
                    $attachment.Load()
                    #call function
                    handleCsv
                }
                else
                {
                    echo "attachment has wrong file format"
                }
            }
        }
        else
        {
            echo "no attachements were found"
        }
    }else
    {
        echo "mail not from today"
    }
}

#end script
exit