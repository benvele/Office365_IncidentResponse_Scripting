# Office365_IncidentResponse_Scripting
This is a main Powershell script used to pull login events from Office 365 and parse them to determine the login location.
It does this by taking the Source IP of the event and running it against a Geo-IP API. It will produce a CSV file with this 
information for further review in Excel or whichever program you wish to inspect with. It also creates PNG pictures of graphs
and maps that show a couple overall statistics of the Office 365 events.
