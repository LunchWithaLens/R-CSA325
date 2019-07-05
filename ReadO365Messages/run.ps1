# BriSmith@Microsoft.com https://blogs.msdn.microsoft.com/brismith
# Code to read O365 Message Cnter posts for specific products then make a Function call with the resultant json

#Setup stuff for the O365 Management Communication API Calls

$password = $env:aad_password | ConvertTo-SecureString -AsPlainText -Force

$Credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $env:aad_username, $password
   
Import-Module "D:\home\site\wwwroot\reado365mcall\Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
   
$adal = "D:\home\site\wwwroot\reado365mcall\Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
[System.Reflection.Assembly]::LoadFrom($adal)
  
$resourceAppIdURI = “https://manage.office.com”
   
$authority = “https://login.windows.net/$env:aadtenant”
   
$authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
$uc = new-object Microsoft.IdentityModel.Clients.ActiveDirectory.UserCredential -ArgumentList $Credential.Username,$Credential.Password

$manageToken = $authContext.AcquireToken($resourceAppIdURI, $env:clientId,$uc)

#Get the products we are interested in
$products = Get-Content 'D:\home\site\wwwroot\reado365mcall\products.json' | Out-String | ConvertFrom-json

###############################################################
# Read service messages posts and get ones of Type MessageCenter
###############################################################

$headers = @{}
$headers.Add('Authorization','Bearer ' + $manageToken.AccessToken)
$headers.Add('Content-Type', "application/json")

$uri = "https://manage.office.com/api/v1.0/" + $env:tenantId + "/ServiceComms/Messages"
$messages = Invoke-WebRequest -Uri $uri -Method Get -Headers $headers -UseBasicParsing
$messagesContent = $messages.Content | ConvertFrom-Json
$messageValue = $messagesContent.Value
ForEach($message in $messageValue){
If($message.MessageType -eq 'MessageCenter'){
ForEach($product in $products){
    If($message.Title -match $product.product){
$task = @{}
$task.Add('id', $message.Id)
$task.Add('title',$message.Id + ' - ' + $message.Title)
$task.Add('categories', $message.ActionType + ', ' + $message.Classification + ', ' + $message.Category)
$task.Add('dueDate', $message.ActionRequiredByDate)
$task.Add('updated', $message.LastUpdatedTime)  
$fullMessage = ''
ForEach($messagePart in $message.Messages){
$fullMessage += $messagePart.MessageText
}
$task.Add('description', $fullMessage)
$task.Add('reference', $message.ExternalLink)
$task.Add('product', $product.product)
$task.Add('bucketId', $product.bucketId)
$task.Add('assignee', $product.assignee)

#Using best practice async via queue storage

$storeAuthContext = New-AzureStorageContext -ConnectionString $env:brismitho365mcr93ec_STORAGE 

$outQueue = Get-AzureStorageQueue –Name 'message-center-to-planner-tasks' -Context $storeAuthContext
if ($outQueue -eq $null) {
    $outQueue = New-AzureStorageQueue –Name 'message-center-to-planner-tasks' -Context $storeAuthContext
}

# Create a new message using a constructor of the CloudQueueMessage class.
$queueMessage = New-Object `
        -TypeName Microsoft.WindowsAzure.Storage.Queue.CloudQueueMessage `
        -ArgumentList (ConvertTo-Json $task)

# Add a new message to the queue.
$outQueue.CloudQueue.AddMessage($queueMessage)
}
}
}
}
