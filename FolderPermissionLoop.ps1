# powershell
# Set the root directory to start from
$rootDirectory = "C:\temp"
$report = "C:\temp\DirectoryPermissions.csv"

# Get all directories in the root directory and its subdirectories
$directories = Get-ChildItem -Path $rootDirectory -Recurse -Directory

# Create an empty list to store the results
$results = @()

# Loop through each directory
foreach ($dir in $directories) {
    # Get user permissions (Access Control List) for the current directory
    $acl = Get-Acl -Path $dir.FullName

    # Loop through each access rule in the ACL
    foreach ($accessRule in $acl.Access) {
        # Extract the user's identity, access type, and file system rights
        $user = $accessRule.IdentityReference.Value
        $accessType = $accessRule.AccessControlType
        $permissions = $accessRule.FileSystemRights

        # Create a custom PowerShell object with the required properties
        $result = New-Object -TypeName PSObject -Property @{
            Directory   = $dir.FullName
            User        = $user
            Access      = $accessType
            Permissions = $permissions
        }

        # Add the custom object to the results list
        $results += $result
    }
}

# Save the results to a CSV file
$results | Export-Csv -Path $report -NoTypeInformation
