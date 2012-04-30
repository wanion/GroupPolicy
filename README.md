# Group Policy PowerShell scripts

Scripts I put together to query Group Policy related information.

## Usage

Load the script into your environment using the dot source operator:

```powershell
. .\Get-GPAudit.ps1
```

Once the script is loaded, you can run the commands like PowerShell Cmdlets.

Example creation of an audit report:

```powershell
# Create an audit
$audit = Get-GPAudit
# Filter down to just the properties you're interested in saving
$filteredaudit = $audit | select displayname, recommendedaction, notes, gpostatus, hasusersection, hascomputersection, linkname, wmifiltername, PermissionsApply, linkpath, wmifilterquery, wmifilterdescription, permissionsother, creationtime, modificationtime, description, id
# Sort results
$sortedaudit = $filteredaudit | sort recommendedaction,notes,linkname,displayname
# Save into file
$sortedaudit | export-csv -path export.csv -notype -force
```

## Scripts

### Get-GPWmiFilter.ps1

Gets WMI filters from Active Directory. Useful as a complement to Backup-GPO to back up WMI filters.

### Get-GPAudit.ps1

**Requires** Group Policy module (Windows Server 2008 R2 with Group Policy Management Console or Windows 7 with Remote Server Administration Tools)

**Depends** Get-GPWmiFilter

Creates a report of group policy objects with recommendations on which could potentially be deleted and reasons why (usually because they're not longer being applied).

Reports the following issues:

* Poorly named (contains underscores or no spaces)
* All settings disabled
* Enabled section is empty
* Both sections are empty
* All links disabled
* Link target has no effect
* No links
* No Apply permissions
* Objects policy applied to have been deleted
* Old test policy (last modification >3 months ago)
* Applies to a single user
* Disabled section is not empty
* ACL contains custom permissions (may mean Apply is denied in some cases)
* At least one link is marked NoOverride (possibly unnecessary, verify if required)

It recommends actions for each policy (rename or delete) based on issues found.