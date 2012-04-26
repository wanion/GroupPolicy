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
$audit = $audit | select displayname, recommendedaction, notes, gpostatus, emptyusersection, emptycomputersection, linkname, wmifiltername, PermissionsApply, linkpath, wmifilterquery, wmifilterdescription, permissionsother, creationtime, modificationtime, description, id
# Sort the way you want
$audit = $audit | sort recommendedaction,notes,linkname,displayname
# Save into file
export-csv -path export.csv -notype
```

## Scripts

### Get-GPWmiFilter.ps1

Gets WMI filters from Active Directory. Useful as a complement to Backup-GPO to back up WMI filters.

### Get-GPAudit.ps1

**Requires** Group Policy module (Windows Server 2008 R2 with Group Policy Management Console or Windows 7 with Remote Server Administration Tools)
**Depends** Get-GPWmiFilter

Creates a report of group policy objects with recommendations on which could potentially be deleted and reasons why (usually because they're not longer being applied).

Examples of detected issues:

* Not linked in any locations.
* Linked to "Group Policies" OU (*who does this?*).
* All settings disabled.
* Computer settings disabled and user settings empty.
* User settings disabled and computer settings empty.
* No user, group, or computer has permission to apply policy.
* ACLs where only trustees with Apply permission are unresolvable SIDs (i.e. policy applies to objects that have now been deleted from AD).
* Policy only applies to a single user.
* Policy contains 'test' in the name and is older than 3 months.

It also recommends renaming policies that are poorly named (currently catches policies that contain no spaces in the name or that contain underscores).