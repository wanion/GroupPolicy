$savepath = '\\tsclient\g\dropbox\work\Current projects\Group policy cleanup\custom-export.csv'

import-module grouppolicy

$wmifilters = @{}

function get-filter([Microsoft.GroupPolicy.Gpo] $gpo) {
  $path = $gpo.Wmifilter.Path
  if ($path -match "\{.+\}")
  {
    $name = $matches[0]
    if (-not $wmifilters.Contains($name))
    {
      Write-Host ("Searching for WMI filter {0}." -f $name)
      $strFilter = "(cn=$name)"
    
      $objDomain = New-Object System.DirectoryServices.DirectoryEntry
      
      $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
      $objSearcher.SearchRoot = $objDomain
      $objSearcher.PageSize = 1000
      $objSearcher.Filter = $strFilter
      $objSearcher.SearchScope = "Subtree"
      $objSearcher.PropertiesToLoad.Add("mswmi-name")
      $objSearcher.PropertiesToLoad.Add("mswmi-parm1")
      $objSearcher.PropertiesToLoad.Add("mswmi-parm2")
      
      $colResults = $objSearcher.FindAll()
      
      if ($colResults.Count -gt 0) {
        foreach ($objResult in $colResults)
        {
          $filter = new-object psobject -property @{
            Guid = $name
            Name = $objResult.Properties["mswmi-name"][0]
            Query = $objResult.Properties["mswmi-parm2"][0]
            Description = $objResult.Properties["mswmi-parm1"][0]
          }
          $wmifilters.add($name,$filter)
          Write-Host ("Found WMI filter {0}." -f $wmifilters[$name].Name)
        }
      }
    }
    return $wmifilters[$name]
  }
}

$allgpo = get-gpo -all
$policies = @()
foreach ($gpo in $allgpo) {
  Write-Host ("Processing {0}..." -f $gpo.DisplayName)
  $wmifilter = get-filter($gpo)
  $report = [xml](get-gporeport -guid $gpo.id -reporttype xml)
  $usersection = $report.gpo.user.versionsysvol -gt 0
  $computersection = $report.gpo.computer.versionsysvol -gt 0
  $permissions = Get-GPPermissions -id $gpo.id -all
  $gpsummary = new-object psobject -property @{
    Id = $gpo.id
    DisplayName = $gpo.DisplayName
    Path = $gpo.Path
    Owner = $gpo.Owner
    DomainName = $gpo.DomainName
    CreationTime = $gpo.CreationTime
    ModificationTime = $gpo.ModificationTime
    User = $gpo.User
    Computer = $gpo.Computer
    GpoStatus = $gpo.GpoStatus
    WmiFilterName = $wmifilter.Name
    WmiFilterQuery = $wmifilter.Query
    WmiFilterDescription = $wmifilter.Description
    LinkName = [string]::join(";", @( $report.gpo.linksto | % { $_.SomName }) )
    LinkPath = [string]::join(";", @( $report.gpo.linksto | % { $_.SomPath }) )
    PermissionsApply = [string]::join(", ", @($permissions | where { $_.permission -eq "GpoApply" } | % { if ($_.Trustee.Name -eq $null) { $_.trustee.sid.value } else { $_.Trustee.Name } }))
    PermissionsOther = [string]::join(", ", @($permissions | where { $_.permission -ne "GpoApply" } | % { "{0}: {1}" -f $_.Trustee.Name, $_.Permission }))
    Description = $gpo.Description
    EmptyUserSection = $usersection -eq $false
    EmptyComputerSection = $computersection -eq $false
  }
  $gpsummary | add-member -membertype scriptproperty -name Notes -value {
    if ($this.gpostatus -eq "AllSettingsDisabled") { return "Disabled already." }
    if (($this.gpostatus -eq "UserSettingsDisabled") -and $this.EmptyComputerSection) { return "Enabled section is empty." }
    if (($this.gpostatus -eq "ComputerSettingsDisabled") -and $this.EmptyUserSection) { return "Enabled section is empty." }
    if ($this.linkname -eq "Group Policies") { return "Linked in a location that has no effect." }
    if ($this.linkname -eq "") { return "Not linked." }
    if ($this.permissionsapply -eq "") { return "No Apply permissions." }
    if ($this.permissionsapply -match '^[S0-9\-,]+$') { return "Objects policy applies to have been deleted." }
    if (($this.displayname -match "test") -and (get-date).addmonths(-3) -gt $this.modificationtime) { return "Old test policy (last modification >3 months ago)." }
    if ($this.permissionsapply -match '^[A-Z]+$') { return "Applies to a single user." }
    if ($this.displayname -match "_") { return "Poorly named (contains underscores)." }
    if (($this.displayname -match " ") -eq $false) { return "Poorly named (no spaces)." }
    if (($this.gpostatus -eq "UserSettingsDisabled") -and $this.EmptyUserSection -ne $true) { return "Disabled section is not empty." }
    if (($this.gpostatus -eq "ComputerSettingsDisabled") -and $this.EmptyComputerSection -ne $true) { return "Disabled section is not empty." }
    return ""
  }
  $gpsummary | add-member -membertype scriptproperty -name RecommendedAction -value {
    if ($this.gpostatus -eq "AllSettingsDisabled") { return "delete" }
    if (($this.gpostatus -eq "UserSettingsDisabled") -and $this.EmptyComputerSection) { return "delete" }
    if (($this.gpostatus -eq "ComputerSettingsDisabled") -and $this.EmptyUserSection) { return "delete" }
    if ($this.linkname -eq "Group policies") { return "delete" }
    if ($this.linkname -eq "") { return "delete" }
    if ($this.permissionsapply -eq "") { return "delete" }
    if ($this.permissionsapply -match '^[S0-9\-,]+$') { return "delete" }
    if (($this.displayname -match "test") -and (get-date).addmonths(-3) -gt $this.modificationtime) { return "delete" }
    if ($this.permissionsapply -match '^[A-Z]+$') { return "delete" }
    if ($this.displayname -match "_") { return "rename" }
    if (($this.displayname -match " ") -eq $false) { return "rename" }
    return ""
  }
  $policies += $gpsummary
}

$policies | select displayname, recommendedaction, notes, gpostatus, emptyusersection, emptycomputersection, linkname, wmifiltername, PermissionsApply, linkpath, wmifilterquery, wmifilterdescription, permissionsother, creationtime, modificationtime, description, id | sort recommendedaction,notes,linkname,displayname | export-csv -force -path $savepath -notype