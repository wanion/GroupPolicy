Import-Module GroupPolicy
. (Join-Path (Split-Path $MyInvocation.MyCommand.Path) Get-GPWMIFilter.ps1)

function Get-GPAudit {
  $GroupPolicyObjects = @(get-gpo -all)
  $GPOCount = $GroupPolicyObjects.Count
  $ProgressCount = 1
  foreach ($gpo in $GroupPolicyObjects) {
    $Progress = [int]($ProgressCount++ / $GPOCount * 100)
    Write-Progress -Activity "Auditing Group Policy objects." -Status ("GPOs processed: {0,2}%" -f $Progress) -CurrentOperation ("Current policy object: {0}" -f $gpo.DisplayName) -PercentComplete $Progress
    if ($gpo.WmiFilter -ne $null) {
      $filterguid = $gpo.wmifilter.path -match "\{.+\}"
      $wmifilter = Get-GPWMIFilter -guid $matches[0]
    }
    $report = [xml](Get-GPOReport -Guid $gpo.id -ReportType XML)
    $HasUserSection = $report.gpo.user.versionsysvol -gt 0
    $HasComputerSection = $report.GPO.Computer.VersionSysvol -gt 0
    $Permissions = Get-GPPermissions -Id $gpo.id -All
    $gpsummary = new-object psobject -Property @{
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
      WmiFilterQuery = $wmifilter.RawQuery
      WmiFilterDescription = $wmifilter.Description
      Links = @( $report.gpo.linksto | % { new-object psobject -property @{
            Name = $_.SOMName
            Path = $_.SOMPath
            Enabled = $_.Enabled
            NoOverride = $_.NoOverride -eq "true"
          }
        }
      )
      PermissionsApply = [string]::join(", ", @($Permissions | where { $_.permission -eq "GpoApply" } | % { if ($_.Trustee.Name -eq $null) { $_.trustee.sid.value } else { $_.Trustee.Name } }))
      PermissionsOther = [string]::join(", ", @($Permissions | where { $_.permission -ne "GpoApply" } | % { "{0}: {1}" -f $_.Trustee.Name, $_.Permission }))
      Description = $gpo.Description
      EmptyUserSection = $HasUserSection -eq $false
      EmptyComputerSection = $HasComputerSection -eq $false
    }
    $gpsummary | add-member -membertype scriptproperty -name LinkName -value {
      return [string]::join(";", @( $this.links | % {
        if ($_.Name -ne $null) {
          $Name = $_.Name
          if ($_.Enabled -eq $false) { $Name += " [Disabled]" }
          if ($_.NoOverride) { $Name += " [NoOverride]" }
          $Name
        }
      }) )
    }
    $gpsummary | add-member -membertype scriptproperty -name LinkPath -value {
      [string]::join(";", @( $this.links | % { $_.Path }) )
    }
    $gpsummary | add-member -membertype scriptproperty -name Notes -value {
      if ($this.gpostatus -eq "AllSettingsDisabled") { return "Disabled already." }
      if (($this.gpostatus -eq "UserSettingsDisabled") -and $this.EmptyComputerSection) { return "Enabled section is empty." }
      if (($this.gpostatus -eq "ComputerSettingsDisabled") -and $this.EmptyUserSection) { return "Enabled section is empty." }
      if ($this.EmptyComputerSection -and $this.EmptyUserSection) { return "Both sections are empty." }
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
    $gpsummary | add-member -membertype scriptproperty -name RecommendedAction -pass -value {
      if ($this.gpostatus -eq "AllSettingsDisabled") { return "delete" }
      if (($this.gpostatus -eq "UserSettingsDisabled") -and $this.EmptyComputerSection) { return "delete" }
      if (($this.gpostatus -eq "ComputerSettingsDisabled") -and $this.EmptyUserSection) { return "delete" }
      if ($this.EmptyComputerSection -and $this.EmptyUserSection) { return "delete" }
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
  }

  Write-Progress -Activity "Auditing Group Policy objects." -Status "GPOs processed %" -Completed
}