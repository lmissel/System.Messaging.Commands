# System.Messaging.Commands

Das PowerShell Module System.Messaging.Commands stellt Funktionen zur Verfügung, für das Herstellen von Verbindungen mit Meldungswarteschlangen im Netzwerk, Überwachen und Verwalten von Meldungswarteschlangen  sowie Senden, Empfangen und Einsehen von Meldungen.

## Beschreibung MSMQ

Message Queuing (MSMQ) ist ein Anwendungsprotokoll von Microsoft, welches Nachrichten-Warteschlangen (Message Queues) zur Verfügung stellt. Unter Windows wird es vom Microsoft Message Queue Server bereitgestellt. MSMQ wird in Software-Anwendungen in serviceorientierten Architekturen eingesetzt.

Mit Hilfe dieser Moduls können Nachrichten an Queues gesendet werden. Ebenso können die Nachrichten aus den Queues abgerufen, und (Receive/Dequeue) werden.

Auch wenn es nicht mehr den Zeitgeist umbedingt immer gerecht werden kann, ist dieses Anwendungsprotokoll für Mikroservices sehr interressant. Vorallem in der Verbindung mit PowerShell können so kleine Services entstehen und durch die MSMQ-Nachrichten Daten an andere Server und Services übermittelt werden.

## Installation

Die Installationhinweise von Message Queuing (MSMQ) entnehmen Sie bitte den Anleitungen von Microsoft. Nachdem Sie Message Queuing (MSMQ) installiert haben, kopieren Sie das Module in eins der PowerShell Module Pfade. 

## Verwendung

```PowerShell

Import-Module System.Messaging.Commands

# Name der Warteschlange
$queueName = ".\private$\vms.mail"

# MSMQ-Message erstellen
$Label = "Mein Label"
$Priority = [System.Messaging.MessagePriority]::Normal
$BodyString = "Mein Text der versendet werden soll..."
$MSMQMessage = New-MSMQMessage -Label $Label -Priority $Priority -Body $BodyString

# Message senden
Send-MSMQMessage -queueName $queueName -MSMQMessage $MSMQMessage -Transaction

# Alle verfügbaren Messages in der Warteschlange anzeigen
Get-MSMQMessageQueue -queueName $queueName

```

## Hinweis
Dieses PowerShell Module wird nicht weiterentwickelt. Bitte verwenden Sie das PowerShell Module "MSMQ".
