<h1>PowerShell Script for connecting to the EWS and writing a mail</h1>
In addition, there is also a filter which you can configure for your own purposes
<br>
<b>For a detailed explanation of the script, please have a look at the <i>Flowchart.png</i> File.</b>
<br><br>
<h2>Things you need to make sure...</h2>
<ul>
  <li>First of all, you should run it on a server so you can create daily/weekly/monthly tasks without difficulties</li>
  <li>Secondly, make sure you put the right information into the variables, otherwise you may send an E-Mail to the wrong person</li>
  <li>Then of course you need to have EWS installed on your server</li>
  <li>You need on software side Exchange Online which is in O365 or use an Exchange Server as on-premise</li>
</ul>
<br>
<h2>What is EWS?</h2>
"Exchange Web Services (EWS) is a cross-platform API that enables applications to access mailbox items such as email messages, meetings, and contacts from Exchange Online, Exchange Online as part of Office 365, or on-premises versions of Exchange starting with Exchange Server 2007. EWS applications can access mailbox items locally or remotely by sending a request in a SOAP-based XML message. The SOAP message is embedded in an HTTP message when sent between the application and the server, which means that as long as your application can post XML through HTTP, it can use EWS to access Exchange."
<br><br>
source: https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/ews-applications-and-the-exchange-architecture
