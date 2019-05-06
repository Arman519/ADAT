import-module smlets
#Get the Incident Class
$IRClass = get-scsmclass -Name System.WorkItem.Incident$
$IncidentImports = Import-CSV "C:\Service Desk Ticketing Tool\ticket_data.csv"
#Set the ID of the Incident
$id = "IR{0}"
 
#Set the title and description of the incident
$title = $IncidentImports.Title
$description = $IncidentImports.Description
 
$impactvalue = $IncidentImports.impact
$urgencyvalue = $IncidentImports.Urgency
#Set the impact and urgency of the incident
$impact = Get-SCSMEnumeration -Name $impactvalue
$urgency = Get-SCSMEnumeration -Name $urgencyvalue
 
#Create a hashtable of the incident values
$incidentHashTable = @{
Id = $id
Title = $title
Description = $description
Impact = $impact
Urgency = $urgency
}
 
#Create the incident
$newIncident = New-SCSMObject $IRClass -PropertyHashtable $incidentHashTable –PassThru
$affectedUserRelClass = Get-SCSMRelationshipClass System.WorkItemAffectedUser$
Write-Host $newIncident

$i = Get-SCSMObject (Get-SCSMClass System.WorkItem.Incident$) -Filter "Id -eq $newIncident"
$assignto = $IncidentImports.Assignedto
$affectedto = $IncidentImports.AffectedUser
$Tierqueto = $IncidentImports.TierQueue
$classificationto = $IncidentImports.Classification
$sourceto = $IncidentImports.Source
$statusto = $IncidentImports.Status
$ticketidto = $i.DisplayName
$ticketlinkto = "http://GLD-win2016-d1:8080/MyRequests/RequestDetails?type=IncidentRequest&id=" + $ticketidto

$UserClass = Get-SCSMClass -name System.Domain.User$
$User = Get-SCSMObject -Class $UserClass -Filter "Username -eq $assignto"
$affectedUser = Get-SCSMObject -Class $UserClass -Filter "Username -eq $affectedto"
$AssignedToUserRelClass  = Get-SCSMRelationshipClass -Name System.WorkItemAssignedToUser$ 
foreach ($incident in $i)
{

New-SCSMRelationshipObject -Relation $AssignedToUserRelClass -Source $incident -Target $User -Bulk  
New-SCSMRelationshipObject -Relation $affectedUserRelClass -Source $incident -Target $affectedUser -Bulk
$i | Set-SCSMObject -Property Classification -Value $IncidentImports.Classification
$i | Set-SCSMObject -Property TierQueue -Value $IncidentImports.TierQueue
$i | Set-SCSMObject -Property Source -Value $IncidentImports.Source
$i | Set-SCSMObject -Property Status -Value $IncidentImports.Status
$impactdisplay = $i.impact.Name
$urgencydisplay = $i.Urgency.Name

}
$info = new-object psobject -property @{Title=$title;Description=$description;Impact=$impactdisplay;Urgency=$urgencydisplay;AssignedTo=$assignto;AffectedUser=$affectedto;TierQueue=$Tierqueto;Classification=$classificationto;Source=$sourceto;Status=$statusto;TicketID=$ticketidto;TicketLink=$ticketlinkto}



$info | select title,description,impact,urgency,assignedto,affecteduser,tierqueue,classification,source,status,ticketid,ticketlink | sort ticketID | Export-csv "C:\Service Desk Ticketing Tool\ticket_data.csv" -Force -NoTypeInformation
