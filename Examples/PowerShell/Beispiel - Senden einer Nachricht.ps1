Import-Module System.Messaging.Commands

# Name der Warteschlange
$queueName = "FormatName:DIRECT=OS:.\private$\vms.mail"

# MSMQ-Message erstellen
$Label = "Mein Label"
$Priority = [System.Messaging.MessagePriority]::Normal
$BodyString = "Mein Text der versendet werden soll..."
$MSMQMessage = New-MSMQMessage -Label $Label -Priority $Priority -BodyString $BodyString

# Message senden
Send-MSMQMessage -queueName $queueName -MSMQMessage $MSMQMessage -Transaction

# Alle verfügbaren Messages in der Warteschlange anzeigen
Get-MSMQMessageQueue -queueName $queueName
