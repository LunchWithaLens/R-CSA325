# R-CSA325
Repo for Microsoft Ready session on O365 Message Center to Planner integration

TL;DR
The Azure fuctions can be used to 
1. Read O365 Comms API, write and the message center messages of interest to an Azure queue
2. Create tasks in a specific bucket in a specific plan assigned to a specific user
Make sure your Azure Function is v1 and uses PowerShell!
Use ADAL v2
Configure a clientId that can read/write Office365 Groups
products.json defines the buckets and assignees, as well as the topics of interest




