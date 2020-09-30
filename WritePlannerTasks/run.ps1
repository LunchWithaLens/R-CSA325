$in = Get-Content $triggerInput -Raw
$messageCenterTask = $in | ConvertFrom-Json
$title = $messageCenterTask.title

# BriSmith@Microsoft.com https://techcommunity.microsoft.com/t5/Planner-Blog/Microsoft-Planner-A-Change-Management-Solution-for-Office-365/ba-p/362360
# Code to read O365 Message Center posts from the message queue and create Planner tasks

# Setup stuff for the Graph API Calls

$password = $env:aad_password | ConvertTo-SecureString -AsPlainText -Force

$Credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $env:aad_username, $password
   
Import-Module "D:\home\site\wwwroot\writeplannertasks\Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
  
$adal = "D:\home\site\wwwroot\writeplannertasks\Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
[System.Reflection.Assembly]::LoadFrom($adal)
  
$resourceAppIdURI = “https://graph.microsoft.com”
   
$authority = “https://login.windows.net/$env:aad_tenant”
   
$authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
$uc = new-object Microsoft.IdentityModel.Clients.ActiveDirectory.UserCredential -ArgumentList $Credential.Username,$Credential.Password

# Get graph token
$graphToken = $authContext.AcquireToken($resourceAppIdURI, $env:clientId,$uc)

$messageCenterPlanId= $env:messageCenterPlanId

#################################################
# Get tasks
# https://docs.microsoft.com/en-us/graph/api/resources/planner-overview?view=graph-rest-1.0
# Shouldn't be an issue with this scenario - but limit on task
# read of around 400 would mean paging required for larger plans
#################################################

$headers = @{}
$headers.Add('Authorization','Bearer ' + $graphToken.AccessToken)
$headers.Add('Content-Type', "application/json")

$uri = "https://graph.microsoft.com/v1.0/planner/plans/" + $messageCenterPlanId + "/tasks"

$messageCenterPlanTasks = Invoke-WebRequest -Uri $uri -Method Get -Headers $headers -UseBasicParsing
$messageCenterPlanTasksContent = $messageCenterPlanTasks.Content | ConvertFrom-Json
$messageCenterPlanTasksValue = $messageCenterPlanTasksContent.value
$messageCenterPlanTasksValue = $messageCenterPlanTasksValue | Sort-Object bucketId, orderHint

#################################################
# Check if the task already exists by bucketId
# I just add tasks I don't already have
#################################################
$taskExists = $FALSE
ForEach($existingTask in $messageCenterPlanTasksValue){
    if(($existingTask.title -match $messageCenterTask.id) -and ($existingTask.bucketId -eq $messageCenterTask.bucketId)){
    $taskExists = $TRUE
    Break
}
}

# Adding the task
if(!$taskExists){
$setTask =@{}
If($messageCenterTask.dueDate){
$setTask.Add("dueDateTime", $messageCenterTask.dueDate)
}
$setTask.Add("orderHint", " !")
# Fixing some weird encoding stuff - probably better way...
$setTask.Add("title", ($messageCenterTask.title -replace "â€™", "'"))
$setTask.Add("planId", $messageCenterPlanId)

# Setting Applied Categories
# Aritrary mapping - other options available
$appliedCategories = @{}
if($messageCenterTask.categories -match 'Action'){
    $appliedCategories.Add("category1",$TRUE)
}
else{$appliedCategories.Add("category1",$FALSE)}
if($messageCenterTask.categories -match 'Plan for Change'){
    $appliedCategories.Add("category2",$TRUE)
}
else{$appliedCategories.Add("category2",$FALSE)}
if($messageCenterTask.categories -match 'Prevent or Fix Issues'){
    $appliedCategories.Add("category3",$TRUE)
}
else{$appliedCategories.Add("category3",$FALSE)}
if($messageCenterTask.categories -match 'Advisory'){
    $appliedCategories.Add("category4",$TRUE)
}
else{$appliedCategories.Add("category4",$FALSE)}
if($messageCenterTask.categories -match 'Awareness'){
    $appliedCategories.Add("category5",$TRUE)
}
else{$appliedCategories.Add("category5",$FALSE)}
if($messageCenterTask.categories -match 'Stay Informed'){
    $appliedCategories.Add("category6",$TRUE)
}
else{$appliedCategories.Add("category6",$FALSE)}

$setTask.Add("appliedCategories",$appliedCategories)

# Set bucket and assignee
# Bucket = product

$setTask.Add("bucketId", $messageCenterTask.bucketId)
$assignmentType = @{}
$assignmentType.Add("@odata.type","#microsoft.graph.plannerAssignment")
$assignmentType.Add("orderHint"," !")
$assignments = @{}
$assignments.Add($messageCenterTask.assignee, $assignmentType)
$setTask.Add("assignments", $assignments)

# Make new task call

$Request = @" 
$($setTask | ConvertTo-Json)
"@

$headers = @{}
$headers.Add('Authorization','Bearer ' + $graphToken.AccessToken)
$headers.Add('Content-Type', "application/json")
$headers.Add('Content-length', + $Request.Length)
$headers.Add('Prefer', "return=representation")
 
$newTask = Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/planner/tasks" -Method Post -Body $Request -Headers $headers -UseBasicParsing
$newTaskContent = $newTask.Content | ConvertFrom-Json
$newTaskId = $newTaskContent.id

# Add task details
# Pull any urls out of the description to add as attachments
$matches = New-Object System.Collections.ArrayList
$matches.clear()
$task.description = $task.description -replace '&amp;', '&'
$regex = 'https?:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=]*)'
# Find all matches in description and add to an array
select-string -Input $messageCenterTask.description -Pattern $regex -AllMatches | % { $_.Matches } | % {     $matches.add($_.Value)}

# Replacing some forbidden characters for odata properties
$externalLink = $messageCenterTask.reference -replace '\.', '%2E'
$externalLink = $externalLink -replace ':', '%3A'
$externalLink = $externalLink -replace '\#', '%23'
$externalLink = $externalLink -replace '\@', '%40'
$messageCenterTask.description = $messageCenterTask.description -replace '[\u201C\u201D]', '"'
$messageCenterTask.description = $messageCenterTask.description -replace "â€œ", '"' 
$messageCenterTask.description = $messageCenterTask.description -replace "â€™", "'"
$messageCenterTask.description = $messageCenterTask.description -replace "â€", '"'
Write-Output $messageCenterTask.description
$setTaskDetails = @{}
$setTaskDetails.Add("description", $messageCenterTask.description)
if(($messageCenterTask.reference) -or ($matches.Count -gt 0)){
$reference = @{}
$reference.Add("@odata.type", "#microsoft.graph.plannerExternalReference")
$reference.Add("alias", "Additional Information")
$reference.Add("type", "Other")
$reference.Add('previewPriority', ' !')
$references = @{}
# Urls with trailing spaces caused failures - hence .trim() below
ForEach($match in $matches){
$match = $match -replace '\.', '%2E'
$match = $match -replace ':', '%3A'
$match = $match -replace '\#', '%23'
$match = $match -replace '\@', '%40'
$references.Add($match.trim(), $reference)
}
if($messageCenterTask.reference){
$references.Add($externalLink.trim(), $reference)
}
$setTaskDetails.Add("references", $references)
$setTaskDetails.Add("previewType", "reference")
}
# The sleep was to ensure the task was persisted before reading back
# Probably a better way via error checking
Start-Sleep -s 2
#Get Current Etag for task details

$uri = "https://graph.microsoft.com/v1.0/planner/tasks/" + $newTaskId + "/details"

$result = Invoke-WebRequest -Uri $uri -Method GET -Headers $headers -UseBasicParsing
$freshEtagTaskContent = $result.Content | ConvertFrom-Json
 
$Request = @"
$($setTaskDetails | ConvertTo-Json)
"@

$headers = @{}
$headers.Add('Authorization','Bearer ' + $graphToken.AccessToken)
$headers.Add('If-Match', $freshEtagTaskContent.'@odata.etag')
$headers.Add('Content-Type', "application/json")
$headers.Add('Content-length', + $Request.Length)
Write-Output $Request
$uri = "https://graph.microsoft.com/v1.0/planner/tasks/" + $newTaskId + "/details"

$result = Invoke-WebRequest -Uri $uri -Method PATCH -Body $Request -Headers $headers -UseBasicParsing
}
Write-Output "PowerShell script processed queue message '$title'"
# Write-Output also useful for debugging the functions
