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

    a.  Open https://portal.azure.com  
    b.  Create a resource  
    c.  __Function App__, give it a name - take defaults - choose a local location  
    d.  We will come back to it later...   
    
2. Get the files you will need (if you are reading this you might already have them)

    a. https://github.com/LunchWithaLens/R-CSA325.git  
    b. If you have Git installed then __git clone https://github.com/LunchWithaLens/R-CSA325.git__  
    
3. Create the target Plan for your Office 365 messages to go into - add a few buckets and members

    a. In the App Launcher of Office 365 (not the Microsoft tenant - as we will not be able to set the right permissions) click __Planner__  
    b. I used a demo tenant, and created a plan and added 4 buckets, 4 members  
    c. From the Url of your Plan - note the groupId and the planId - you will need them later  
    
4. Identify the IDs of the buckets using Graph Explorer

    a. https://developer.microsoft.com/en-us/graph/graph-explorer  
    b. Sign in with Microsoft – to see your Plan  
    c. https://graph.microsoft.com/v1.0/me/Planner/Plans  
    d. Modify permissions (add Group Read/Write All)  
    e. Give consent (If you have used Graph against your tenant before then this and the previous step may not be needed)  
    f. https://graph.microsoft.com/v1.0/me/Planner/Plans  
    g. https://graph.microsoft.com/v1.0/Planner/Plans/<__planId__>/Buckets  
    h. Note the Id for each of your buckets  
    
5. Identify your members using Graph Explorer

    a. https://graph.microsoft.com/v1.0/Groups/<__GroupId__>/Members  
    b. Note the IDs of your members  
    
6. Construct your products.json.  The product may or may not match the bucket, but obviously should be related.  I have used some product names and also new and updated to get specific types of messages.  
A message may be copied to more than one bucket if it matches more than one product.  The bucket IDs come from step 4, and the assignee IDs from step 5.

    
    >[  
    >>    {  
    >>>        "product": “New",  
    >>>        "bucketId": "yLZpnXLpJUmGRECbFiogsQgAEGcb",   
    >>>        "assignee": "2545703b-333d-4688-ad69-4d2b6f2ceafe"  
    >>    },  
    >>    {  
    >>>        "product": “Updated",     
    >>>        "bucketId": "LZ53WX0A-USEY1CuVBv0RggAC8_Y",   
    >>>        "assignee": "a48e6530-3ceb-47b6-889d-53613159597d"  
    >>    },…  

7. Get a __clientId__ for your application, from your __Office365 Admin Center, Azure Active Directory__

    a. Navigate to __Azure Active Directory Admin Center__ (__App Launcher, Admin, Show All, Admin Centers, Azure Active Directory__)  
    b. Select __App registrations__    
    c. Name the App and register - make a note of the Application (client) ID  
    d. __API Permissions__, __Add a permisison__, then add Office 365 Management APIs – Delegated permissions, ServiceHealth.Read   and 
    Microsoft Graph, Delegated Permissions, Group.ReadWrite.All  
    e. __Grant admin consent for Contoso__ (or to whoever your tenant is)  
    f. Authentication – Advanced settings – Default client type. Set to Yes for Treat application as a public client  
    
8. Your Azure function should be ready now - so back to your Azure portal and find your Function App

    a. __IMPORTANT!__  __Platform features__, __Function App settings__ and change Runtime version to 1  
    b. Add a function by clicking the __+__, choose __create your own custom function__    
    c. Enable Experimental Language Support to light up PowerShell, then choose PowerShell in the Timer trigger language list  
    d. Leave __language__ as PowerShell, set the __name__ to ReadO365Messages, and set the __schedule__ to 0 0 16 * * * (4pm UTC daily)  
    
9. Configuring the ReadO365Messages function

    a. Under __Integrate__ click __+ New Output__, and choose __Azure Queue Storage__, then __Select__  
    b. Change the __Queue name__ to message-center-to-planner-tasks and click __Save__, then once saved click __Go__ next to __Create a new function triggered by this output__  
    c. Enable __Experimental Language Support__ again, then choose __PowerShell__ in the __Queue trigger__ language list  
    d. __Name__ the function WritePlannerTasks, and change the __Queue name__ to message-center-to-planner-tasks, then click __Create__  
    
10. Put our PowerShell code and required DLL in place

    a. Copy the two run.ps1 samples in place of the default code - the directories in the Github repro match the function names.  _If you didn't use my suggested names for the functions then the paths in the code will need to be edited_  
    b. Upload the DLL - Microsoft.IdentityModel.Clients.ActiveDirectory.dll - by clicking the __View files__ under each function then selecting Upload - upload to each function  
    c. Upload your _products.json_ to the ReadO365Messages function (Don't use mine - the IDs will all be wrong)
    
11. Add your variables

    a. At the root of the Function App, Configuration, and add new application settings for:
    
    * aad_tenant
    * aad_username
    * aad_password
    * clientId
    * tenantId
    * messageCenterPlanId

    Don't forget to Save
    
At this point the ReadO365Messages function can be run.  It may fail first time as it appears the creating of the queue happens, but is still not found to write to.  Running again should then work, and the schedule should run from there on.  

Last step is to add the text to the coloured labels.  This can be done on any of the tasks and is then set for the plan
* Action
* Plan for Change
* Prevent or Fix Issues
* Advisory
* Awareness
* Stay Informed






