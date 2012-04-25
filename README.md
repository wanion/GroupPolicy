# Group Policy scripts

A couple of scripts I put together for my own use. Both could use a little tidying but shouldn't be difficult to adapt as they are.

## export-wmifilters.ps1

Saves a text file containing all the WMI filters in the current domain. Useful complement to backup-gpo cmdlet.

You'll want to modify the path it saves to:

```powershell
$path = new-item -path ("{0}\{1}" -f (Split-Path -parent $MyInvocation.MyCommand.Definition), "GPO_" + (get-date -uformat "%Y-%m-%d")) -Type directory -force
```

## new-gpreport.ps1

> Requires: Group Policy module

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