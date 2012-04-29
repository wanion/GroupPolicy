Import-Module GroupPolicy
. (Join-Path (Split-Path $MyInvocation.MyCommand.Path) Get-GPWMIFilter.ps1)

function Get-GPAudit {
  $GroupPolicyObjects = @(get-gpo -all)
  $GPOCount = $GroupPolicyObjects.Count
  $ProgressCount = 1
  foreach ($gpo in $GroupPolicyObjects) {
    $Progress = [int]($ProgressCount++ / $GPOCount * 100)
    Write-Progress -Activity "Auditing Group Policy objects." -Status ("GPOs processed: {0,2}%" -f $Progress) -CurrentOperation ("Current policy object: {0}" -f $gpo.DisplayName) -PercentComplete $Progress
    $wmifilter = new-object psobject
    if ($gpo.wmifilter.path -match "\{.+\}") {
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
      RecommendedAction = ""
      Notes = @() # Placeholder array
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

    # Determine recommended action for this policy
    if ($gpsummary.displayname -match "_") {
      $gpsummary.Notes += "Poorly named (contains underscores)"
      $gpsummary.RecommendedAction= "rename"
    }
    if (($gpsummary.displayname -match " ") -eq $false) {
      $gpsummary.Notes += "Poorly named (no spaces)"
      $gpsummary.RecommendedAction= "rename"
    }
    if ($gpsummary.gpostatus -eq "AllSettingsDisabled") { 
      $gpsummary.Notes +=  "All settings disabled"
      $gpsummary.RecommendedAction= "delete"
    }
    if ((($gpsummary.gpostatus -eq "UserSettingsDisabled") -and $gpsummary.EmptyComputerSection) -or (($gpsummary.gpostatus -eq "ComputerSettingsDisabled") -and $gpsummary.EmptyUserSection)) {
      $gpsummary.Notes += "Enabled section is empty"
      $gpsummary.RecommendedAction= "delete"
    }
    if ($gpsummary.EmptyComputerSection -and $gpsummary.EmptyUserSection) {
      $gpsummary.Notes += "Both sections are empty"
      $gpsummary.RecommendedAction= "delete"
    }
    if (($gpsummary.linkname -match "^Group Policies") -and ($gpsummary.Links.Count -eq 1)) {
      $gpsummary.Notes += "Link target has no effect"
      $gpsummary.RecommendedAction= "delete"
    }
    if ($gpsummary.linkname -eq "") {
      $gpsummary.Notes += "No links"
      $gpsummary.RecommendedAction= "delete"
    }
    if ($gpsummary.permissionsapply -eq "") {
      $gpsummary.Notes += "No Apply permissions"
      $gpsummary.RecommendedAction= "delete"
    }
    if ($gpsummary.permissionsapply -match '^[S0-9\-,]+$') {
       $gpsummary.Notes +=  "Objects policy applied to have been deleted"
       $gpsummary.RecommendedAction= "delete"
    }
    if (($gpsummary.displayname -match "test") -and (get-date).addmonths(-3) -gt $gpsummary.modificationtime) { 
      $gpsummary.Notes += "Old test policy (last modification >3 months ago)"
      $gpsummary.RecommendedAction= "delete"
    }
    if ($gpsummary.permissionsapply -match '^[A-Z]+$') {
      $gpsummary.Notes += "Applies to a single user"
      $gpsummary.RecommendedAction= "delete"
    }
    if ((($gpsummary.gpostatus -eq "UserSettingsDisabled") -and $gpsummary.EmptyUserSection -ne $true) -or (($gpsummary.gpostatus -eq "ComputerSettingsDisabled") -and $gpsummary.EmptyComputerSection -ne $true)) {
      $gpsummary.Notes += "Disabled section is not empty"
    }

    $gpsummary.Notes = [string]::join("; ", $gpsummary.Notes) # Change notes from array to string

    $gpsummary
  }

  Write-Progress -Activity "Auditing Group Policy objects." -Status "GPOs processed %" -Completed
}