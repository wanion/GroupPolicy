function Get-GpoLinks {
  $searcher = new-object system.directoryservices.directorysearcher
  $searcher.filter = "(|(objectclass=organizationalunit)(objectclass=domain))"
  $searcher.propertiestoload.add("distinguishedname") | out-null
  $links = $searcher.findall() | % { get-gpinheritance $_.properties.distinguishedname }

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