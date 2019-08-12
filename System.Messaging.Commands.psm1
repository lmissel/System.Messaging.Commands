#**************************************************
#
# Module: System.Messaging.Commands
#
# Written by: lmissel
#
# Beschreibung:
# ------------
# Der System.Messaging-Namespace stellt Klassen für folgende Aufgaben zur Verfügung: 
# Herstellen von Verbindungen mit Meldungswarteschlangen im Netzwerk, Überwachen und Verwalten von Meldungswarteschlangen 
# im Netzwerk sowie Senden, Empfangen und Einsehen von Meldungen.
#
# Hinweis:
# -----
# Dieses Module wird nicht weiterentwickelt. Bitte verwenden Sie das PowerShell Module "MSMQ".
# LINK - https://docs.microsoft.com/en-us/powershell/module/msmq/?view=win10-ps
#
#**************************************************

[Reflection.Assembly]::LoadWithPartialName("System.Messaging") | Out-Null

#**************************************************
# MSMQ - MessageQueue
#**************************************************

function New-MSMQMessageQueue
{
    param(
        [String] $queuePath
    )

    Begin
    {
    }

    Process
    {
        try
        {
            if(!([System.Messaging.MessageQueue]::Exists($queuePath)))
            {
               [System.Messaging.MessageQueue]::Create($queuePath);
            }
            else
            {
                [Console]::WriteLine($queuePath + " already exists.")
            }
        }
        catch [System.Messaging.MessageQueueException]
        {
            [Console]::WriteLine($_.Exception.Message)
        }
    }

    End

    {
    }
}

function Get-MSMQPublicMessageQueue
{
    param(
        [Switch] $All
    )

    Begin
    {
    }

    Process
    {
        if ($PSCmdlet.ParameterSetName -eq 'All')
        {
            #Get a list of public queues.
		    [System.Messaging.MessageQueue[]] $QueueList = [System.Messaging.MessageQueue]::GetPublicQueues();
            return $QueueList
        }
    }

    End
    {
    }
}

function Get-MSMQMessageQueue
{
    param(
        [String] $queuePath
    )

    Begin
    {
    }

    Process
    {
        try
        {
            if(!([System.Messaging.MessageQueue]::Exists($queuePath)))
            {
               [Console]::WriteLine("The messageQueue does not existed.")
            }
            else
            {
                [System.Messaging.MessageQueue] $MessageQueue = [System.Messaging.MessageQueue]::New($queuePath)
                return $MessageQueue
            }
        }
        catch [System.Messaging.MessageQueueException]
        {
            if($_.Exception.MessageQueueErrorCode -eq [System.Messaging.MessageQueueErrorCode]::AccessDenied)
			{
				[Console]::WriteLine("Access is denied. Queue might be a system queue.");
			}
            else
            {
                [Console]::WriteLine($_.Exception.Message)
            }
        }
    }

    End
    {
    }
}

function Remove-MSMQMessageQueue
{
    param(
        [String] $queuePath
    )

    Begin
    {
    }

    Process
    {
        try
        {
            if(!([System.Messaging.MessageQueue]::Exists($queuePath)))
            {
               [Console]::WriteLine("The messageQueue does not existed.")
            }
            else
            {
                [System.Messaging.MessageQueue]::Delete($queuePath);
            }
        }
        catch [System.Messaging.MessageQueueException]
        {
            if($_.Exception.MessageQueueErrorCode -eq [System.Messaging.MessageQueueErrorCode]::AccessDenied)
			{
				[Console]::WriteLine("Access is denied. Queue might be a system queue.");
			}
            else
            {
                [Console]::WriteLine($_.Exception.Message)
            }
        }
    }

    End
    {
    }
}

function Register-ReceiveCompletedEvent
{
    param(
        [System.Messaging.MessageQueue] $MessageQueue
    )

    Register-ObjectEvent -InputObject $MessageQueue -EventName "ReceiveCompleted" -Action {
        (New-Event -SourceIdentifier "ReceiveCompleted" -Sender $args[0] –EventArguments $args[1])
    }
}

function Register-PeekCompletedEvent
{
    param(
        [System.Messaging.MessageQueue] $MessageQueue
    )

    Register-ObjectEvent -InputObject $MessageQueue -EventName "PeekCompleted" -Action {
        (New-Event -SourceIdentifier "PeekCompleted" -Sender $args[0] –EventArguments $args[1])
    }
}

#**************************************************
# MSMQ - Message
#**************************************************

function New-MSMQMessage
{
    [CmdletBinding(DefaultParameterSetName='Empty', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  HelpUri = 'http://www.microsoft.com/',
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([System.Messaging.Message])]
    param(
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Default')]
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Object')]
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='IMessageFormatter')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [String] $Label,

        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='Default')]
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='Object')]
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='IMessageFormatter')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [System.Messaging.MessagePriority] $Priority,

        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=2,
                   ParameterSetName='Object')]
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=2,
                   ParameterSetName='IMessageFormatter')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Object] $body,

        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=3,
                   ParameterSetName='IMessageFormatter')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [System.Messaging.IMessageFormatter] $formatter
    )

    Begin
    {
    }

    Process
    {
        try
        {
            if ($PSCmdlet.ParameterSetName -eq 'Empty')
            {
                [System.Messaging.Message] $Message = [System.Messaging.Message]::new()

                return $Message
            }

            if ($PSCmdlet.ParameterSetName -eq 'Default')
            {
                [System.Messaging.Message] $Message = [System.Messaging.Message]::new()
                $Message.Label = $Label
                $Message.Priority = $Priority

                return $Message
            }

            if ($PSCmdlet.ParameterSetName -eq 'Object')
            {
                [System.Messaging.Message] $Message = [System.Messaging.Message]::new($body)
                $Message.Label = $Label
                $Message.Priority = $Priority

                return $Message
            }

            if ($PSCmdlet.ParameterSetName -eq 'IMessageFormatter')
            {
                [System.Messaging.Message] $Message = [System.Messaging.Message]::new($body, $formatter)
                $Message.Label = $Label
                $Message.Priority = $Priority

                return $Message
            }
        }
        catch
        {
            [Console]::WriteLine($_.Exception.Message)
        }
    }

    End
    {
    }
}

function Get-MSMQMessage
{
    [CmdletBinding(DefaultParameterSetName='All', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  HelpUri = 'http://www.microsoft.com/',
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([System.Messaging.Message])]
    param(
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='All')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [String] $queueName,
        
        # GetAllMessages()
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='All')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Switch] $All,
        
        # Set the formatter
        [Type[]] $XmlMessageFormatter
    )

    Begin
    {
    }

    Process
    {
        # Connect to a queue.
        [System.Messaging.MessageQueue] $queue = [System.Messaging.MessageQueue]::new($queueName)

        if ($XmlMessageFormatter)
        {
            $queue.Formatter = [System.Messaging.XmlMessageFormatter]::new($XmlMessageFormatter)
        }

        try
		{
            if ($PSCmdlet.ParameterSetName -eq 'All')
            {
                # Populate an array with copies of all the messages in the queue.
                [System.Messaging.Message[]] $msgs = $queue.GetAllMessages()

                # Loop through the messages.
                foreach ($msg in $msgs)
                {
                    # Display each message.
                    $msg
                }
            }
        }
        catch [MessageQueueException]
        {
            # Handle Message Queuing exceptions.
            [Console]::WriteLine($_.Exception.Message)
        }
        catch [InvalidOperationException]
        {
            # Handle invalid serialization format.
            [Console]::WriteLine($_.Exception.Message)
        }
        catch
        {
            #Catch other exceptions as necessary.
            [Console]::WriteLine($_.Exception.Message)
        }
    }

    End
    {
    }
}

function Send-MSMQMessage
{
    param(
        [String] $queueName,
        [System.Messaging.Message] $MSMQMessage,
        [Switch] $Transaction
    )

    try
    {
        # Connect to a queue.
        $queue = new-object System.Messaging.MessageQueue $queueName

        # Send the message to the queue.
        if ($Transaction)
        {
            $MessageQueueTransaction = new-object System.Messaging.MessageQueueTransaction
            $MessageQueueTransaction.Begin()
            $queue.Send($MSMQMessage)
            $MessageQueueTransaction.Commit()
        }
        else
        {
            $queue.Send($MSMQMessage)
        }
    }
    catch [System.ArgumentException]
    {
        [Console]::WriteLine($_.Exception.Message)
    }
}

function Receive-MSMQMessage
{
    [CmdletBinding(DefaultParameterSetName='All', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  HelpUri = 'http://www.microsoft.com/',
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([System.Messaging.Message])]
    param(
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='All')]
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Id')]
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Next')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [String] $queueName,

        # ReceiveById($Id)
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='Id')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [String] $Id, 

        # Receive()
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='Next')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Switch] $Next, 

        # GetAllMessages(), Receive()
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='All')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Switch] $All,
        
        # Set the formatter
        [Type[]] $XmlMessageFormatter
    )

    Begin
    {
    }

    Process
    {
        [System.Messaging.MessageQueue] $queue = [System.Messaging.MessageQueue]::new($queueName)

        if ($XmlMessageFormatter)
        {
            $queue.Formatter = [System.Messaging.XmlMessageFormatter]::new($XmlMessageFormatter)
        }

        try
		{
            # GetAllMessages(), Receive()
            if ($PSCmdlet.ParameterSetName -eq 'All')
            {
                $msgs = $queue.GetAllMessages()
                foreach ($msg in $msgs)
                {
                    $queue.ReceiveById($msg.Id)
                }
            }

            # ReceiveById($Id)
            if ($PSCmdlet.ParameterSetName -eq 'Id')
            {
                $queue.ReceiveById($Id)
            }

            # Receive()
            if ($PSCmdlet.ParameterSetName -eq 'Next')
            {
                $queue.Receive()
            }
        }
        catch [MessageQueueException]
        {
            # Handle Message Queuing exceptions.
            [Console]::WriteLine($_.Exception.Message)
        }
        catch [InvalidOperationException]
        {
            # Handle invalid serialization format.
            [Console]::WriteLine($_.Exception.Message)
        }
        catch
        {
            #Catch other exceptions as necessary.
            [Console]::WriteLine($_.Exception.Message)
        }
    }

    End
    {
    }
}

function Receive-MSMQMessageAsync
{
    [CmdletBinding(DefaultParameterSetName='All', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  HelpUri = 'http://www.microsoft.com/',
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([System.Messaging.Message])]
    param(
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='All')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [String] $queueName,
        
        # Set the formatter
        [Type[]] $XmlMessageFormatter
    )

    Begin
    {
    }

    Process
    {
        # Connect to a queue.
        [System.Messaging.MessageQueue] $queue = [System.Messaging.MessageQueue]::new($queueName)

        try
		{
            # Starten der asynchronen Abfrage
            $asyncResult = $queue.BeginReceive()

            # Hier kann noch weiterer Code stehen...

            # Beenden der asynchronen Abfrage 
            [System.Messaging.Message] $Message = $queue.EndReceive($asyncResult)
            
            # Ergebnis ausgeben
            return $Message
        }
        catch [System.Messaging.MessageQueueException]
        {
            # Handle Message Queuing exceptions.
            [Console]::WriteLine($_.Exception.Message)
        }
        catch [InvalidOperationException]
        {
            # Handle invalid serialization format.
            [Console]::WriteLine($_.Exception.Message)
        }
        catch
        {
            #Catch other exceptions as necessary.
            [Console]::WriteLine($_.Exception.Message)
        }
    }

    End
    {
    }
}

function Wait-MSMQMessage
{
    [CmdletBinding(DefaultParameterSetName='All', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  HelpUri = 'http://www.microsoft.com/',
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([System.Messaging.Message])]
    param(
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='All')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [String] $queueName,

        # Set the formatter
        [Type[]] $XmlMessageFormatter
    )

    Begin
    {
    }

    Process
    {
        [System.Messaging.MessageQueue] $queue = [System.Messaging.MessageQueue]::new($queueName)
        
        if ($XmlMessageFormatter)
        {
            $queue.Formatter = [System.Messaging.XmlMessageFormatter]::new($XmlMessageFormatter)
        }

        try
		{
            $queue.MessageReadPropertyFilter.ClearAll()
            [Void] $queue.Peek()
            [Console]::WriteLine("A message has arrived in the queue.")
        }
        catch [MessageQueueException]
        {
            # Handle Message Queuing exceptions.
            [Console]::WriteLine($_.Exception.Message)
        }
        catch [InvalidOperationException]
        {
            # Handle invalid serialization format.
            [Console]::WriteLine($_.Exception.Message)
        }
        catch
        {
            #Catch other exceptions as necessary.
            [Console]::WriteLine($_.Exception.Message)
        }
    }

    End
    {
    }
}

function Wait-MSMQMessageAsync
{
    [CmdletBinding(DefaultParameterSetName='All', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  HelpUri = 'http://www.microsoft.com/',
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([System.Messaging.Message])]
    param(
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='All')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [String] $queueName,

        # Set the formatter
        [Type[]] $XmlMessageFormatter
    )

    Begin
    {
    }

    Process
    {
        [System.Messaging.MessageQueue] $queue = [System.Messaging.MessageQueue]::new($queueName)
        
        if ($XmlMessageFormatter)
        {
            $queue.Formatter = [System.Messaging.XmlMessageFormatter]::new($XmlMessageFormatter)
        }

        try
		{
            $queue.MessageReadPropertyFilter.ClearAll()
            $asyncResult = $queue.BeginPeek()

            [System.Messaging.Message] $Message = $queue.EndPeek($asyncResult)
            [Console]::WriteLine("A message [$($Message.Id)] has arrived in the queue.")
        }
        catch [MessageQueueException]
        {
            # Handle Message Queuing exceptions.
            [Console]::WriteLine($_.Exception.Message)
        }
        catch [InvalidOperationException]
        {
            # Handle invalid serialization format.
            [Console]::WriteLine($_.Exception.Message)
        }
        catch
        {
            #Catch other exceptions as necessary.
            [Console]::WriteLine($_.Exception.Message)
        }
    }

    End
    {
    }
}

function Read-MSMQMessage
{
    [CmdletBinding(DefaultParameterSetName='Default', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  HelpUri = 'http://www.microsoft.com/',
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([Object])]
    param(
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Default')]
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='IMessageFormatter')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [System.Messaging.Message] $Message,

        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='IMessageFormatter')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [System.Messaging.IMessageFormatter] $Formatter
    )

    Begin
    {
    }

    Process
    {
        if ($PSCmdlet.ParameterSetName -eq 'Default')
        {
            [System.IO.StreamReader]$StreamReader = [System.IO.StreamReader]::new($Message.BodyStream)
            $object = $StreamReader.ReadToEnd()
            return $object
        }

        if ($PSCmdlet.ParameterSetName -eq 'IMessageFormatter')
        {
            [System.IO.StreamReader]$StreamReader = [System.IO.StreamReader]::new($Message.BodyStream)
            $Message.Formatter = $Formatter
            $object = $StreamReader.ReadToEnd()
            return $object
        }
    }

    End
    {
    }
}