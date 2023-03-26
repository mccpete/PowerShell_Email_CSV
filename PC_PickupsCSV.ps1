#Pickups with csv

$OL = New-Object -ComObject outlook.application

 

$csvfilepath = "C:\Users\Name\Documents\data.csv"

 

$PCData  = Import-Csv -Path $csvfilepath -Delimiter ','

 

 

foreach($PC in $PCData){

   

     $computer_name = $PC.PCName

     $serial = $PC.Serial

     $fromuser = $PC.FromUser

     $man_email = $PC.ManagerEmail

     $man_fname = $PC.ManagerFname

 

     <#write-host $computer_name

     write-host $serial

     write-host $fromuser

     write-host $man_email

     write-host $man_fname

     #>

 

     $mItem = $OL.CreateItem("olMailItem")

    $SendTO = $man_email

 

    $mItem.To = $SendTo

    $serial_number = $serial

 

    $mItem.Subject = "PC Pickup $serial"

 

    $personfirst_name = $man_fname

 

    $pc_name = $computer_name

 

    $from_user = $fromuser

 

$signature = "

 

Thanks,

 

Your Email Signature here"

 

$email_body = "Hello $personfirst_name,

 

I have a ticket to reclaim

PC Name: $pc_name

Serial Number: $serial_number

From User: $from_user

Let me know how to connect with you to best retrieve this asset.

I am located at the Glendale HQ office.

 

$signature

"

 

write-host $email_body

 

$mItem.Body = $email_body

 

 

 

$mItem.Send()

 

}

 
