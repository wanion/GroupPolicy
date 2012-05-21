function Get-GpoLinks {
  $links = adfind -default -f "|(objectclass=organizationalunit)(objectclass=domain)" -list distinguishedname | % { get-gpinheritance $_ }

  foreach ($link in $links) {
    new-object psobject -property @{
      Name = $link.Name
      Path = $link.Path
      ContainerType = $link.ContainerType
      GpoInheritanceBlocked = $link.GpoInheritanceBlocked
      GpoLinksEnabled = [string]::Join("; ", @($link.GpoLinks | where { $_.enabled -eq $true } | % {
        $LinkName = $_.DisplayName
        if ($_.enforced) { $LinkName += " [enforced]" }
        $LinkName
      }))
      GpoLinksDisabled = [string]::Join("; ", @($link.GpoLinks | where { $_.enabled -eq $false } | % {
        $LinkName = $_.DisplayName
        if ($_.enforced) { $LinkName += " [enforced]" }
        $LinkName
      }))
      InheritedGpoLinks = [string]::Join("; ", @($link.InheritedGpoLinks | % {
        $LinkName = $_.DisplayName
        if ($_.enforced) { $LinkName += " [enforced]" }
        if (-not $_.enabled) { $LinkName += " [disabled]" }
        $LinkName
      }))
    }
  }
}