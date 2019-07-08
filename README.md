# R-CSA325
## Repo for Microsoft Ready session on O365 Message Center to Planner integration

TL;DR
The Azure fuctions can be used to 
1. Read O365 Comms API, write the message center messages of interest to an Azure queue
2. Create new messages as tasks in a specific bucket in a specific plan assigned to a specific user

Make sure your Azure Function is v1 and uses PowerShell!
Use ADAL v2.28 or 2.29
Configure a clientId that can read/write Office365 Groups, and read O365 Management APIs, ServiceHealth
products.json defines the buckets and assignees, as well as the topics of interest

Detailed steps:
1. Get the Function App started - it only takes a few minutes but will save a wait later
    a. Open https://portal.azure.com
    b. Create a resource
    c. Function App, give it a name - take defaults - choose a local location
    d. We will come back to it later...
2. Get the files you will need (if you are reading this you might already have them)
    a. https://github.com/LunchWithaLens/R-CSA325.git
    b. If you have Git installed then git clone https://github.com/LunchWithaLens/R-CSA325.git
3. Create the target Plan for your Office 365 messages to go into - add a few buckets and members
    a. In the App Launcher of Office 365 (not the Microsoft tenant - as we will not be able to set the right permissions) click Planner
    b. I used a demo tenant, and created a plan and added 4 buckets, 4 members.  
    c. From the Url of your Plan - note the groupId and the planId - you will need them later
4. Identify the IDs of the buckets using Graph Explorer
    a. https://developer.microsoft.com/en-us/graph/graph-explorer
    b. Sign in with Microsoft – to see your Plan
    c. https://graph.microsoft.com/v1.0/me/Planner/Plans
    d. Modify permissions (add Group Read/Write All)
    e. Give consent
    f. https://graph.microsoft.com/v1.0/me/Planner/Plans
    g. https://graph.microsoft.com/v1.0/Planner/Plans/<planId>/Buckets
    h. Note the Id for each of your buckets
5. Identify your members using Graph Explorer
    a. Plan members = Group members
    b. https://graph.microsoft.com/v1.0/Groups/<GroupId>/Members
    c. Note the IDs of your members
6. Construct your products.json.  The product may or may not match the bucket, but obviously should be related.  I have used some product names and also new and updated to get specific types of messages.  A message may be copied to more than one bucket if it matches more than one product.  The bucket IDs come from step 4, and the assignee IDs from step 5.

    [
        {
            "product": “New", 
            "bucketId": " yLZpnXLpJUmGRECbFiogsQgAEGcb ", 
            "assignee": " 2545703b-333d-4688-ad69-4d2b6f2ceafe "
        },
        {
            "product": “Updated", 
            "bucketId": " LZ53WX0A-USEY1CuVBv0RggAC8_Y ", 
            "assignee": " a48e6530-3ceb-47b6-889d-53613159597d "
        },…









