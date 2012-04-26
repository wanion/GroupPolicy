function Get-GPWMIFilter() {
  # Set up query
  $searcher = New-Object System.DirectoryServices.DirectorySearcher
  $searcher.Filter = "(objectclass=msWMI-Som)"

  # Select properties to retrieve
  $properties = @("mswmi-name", "mswmi-parm1", "mswmi-parm2", "cn")
  $properties | % { $searcher.PropertiesToLoad.Add($_) | out-null }

  # Loop through results and dump to file
  foreach ($result in $searcher.FindAll())
  {
    # Split query details into individual objects
    $result.Properties["mswmi-parm2"][0] -match "^([0-9]+);(.+)" | out-null
    $QueryCount = $matches[1]
    $Queries = @()
    $RawQuery = $matches[2]
    $SplitQuery = $RawQuery.Split(";")
    for ($i = 1; $i -le $QueryCount; $i++) {
      $Queries += new-object psobject -property @{
        Namespace = $SplitQuery[$i*6-2]
        Query = $SplitQuery[$i*6-1]
      }
    }

    # Create objects for each filter
    new-object psobject -property @{
      Guid = [guid]$result.Properties["cn"][0]
      Name = $result.Properties["mswmi-name"][0]
      Description = $result.Properties["mswmi-parm1"][0]
      RawQuery = $result.Properties["mswmi-parm2"][0]
      Query = $Queries
    }
  }
}