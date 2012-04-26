function export-wmifilters([string] $path) {
  # Remove existing export if it exists, don't want to append to it
  remove-item -force -path ($path + "\wmi filters.txt") -erroraction:silentlycontinue

  # Set up query
  $searcher = New-Object System.DirectoryServices.DirectorySearcher
  $searcher.Filter = "(objectclass=msWMI-Som)"

  # Select properties to retrieve
  $properties = @("mswmi-name", "mswmi-parm1", "mswmi-parm2", "cn")
  $properties | % { $searcher.PropertiesToLoad.Add($_) | out-null }

  # Loop through results and dump to file
  $counter = 1
  foreach ($result in $searcher.FindAll())
  {
    @'
Policy {0}
Name: {1}
Description: {2}
Guid: {3}
Query: {4}

'@ -f $counter++, $result.Properties["mswmi-name"][0], $result.Properties["mswmi-parm1"][0], $result.Properties["cn"][0], $result.Properties["mswmi-parm2"][0] | out-file -filepath ($path + "\wmi filters.txt") -append
  }
}

$path = new-item -path ("{0}\{1}" -f (Split-Path -parent $MyInvocation.MyCommand.Definition), "GPO_" + (get-date -uformat "%Y-%m-%d")) -Type directory -force
export-wmifilters($path.fullname)