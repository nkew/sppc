# Install the PnP PowerShell module if not already installed
# Install-Module -Name "PnP.PowerShell" -Force

param (
    [string]$SiteURL = "https://xxx.sharepoint.com/sites/ProductDev",
    [string]$ReportFile = "C:\Temp\PermissionRpt.csv",
    [int]$BatchSize = 500,
    [switch]$FoldersOnly
)

Import-Module "PnP.PowerShell"

# Helper function to execute the query with retry logic
Function Execute-QueryWithRetry([scriptblock]$QueryBlock, [int]$RetryCount = 10, [int]$InitialDelay = 5) {
    $attempt = 0
    $success = $false
    $delay = $InitialDelay

    while (-not $success -and $attempt -lt $RetryCount) {
        Try {
            Write-Host -f Yellow "Attempting request to SharePoint... Attempt $($attempt + 1) of $RetryCount"
            & $QueryBlock
            $success = $true
            Write-Host -f Green "Request successful."
        }
        Catch {
            $attempt++
            if ($_.Exception -and $_.Exception.Response -and $_.Exception.Response.StatusCode -eq 429) {
                Write-Host -f Red "Request throttled. Retrying in $delay seconds... (Attempt $attempt of $RetryCount)"
                Start-Sleep -Seconds $delay
                $delay *= 2  # Exponential backoff
            }
            Else {
                Write-Host -f Red "Error encountered: $($_.Exception.Message)"
                throw $_.Exception  # Re-throw the exception if it's not a 429 error
            }
        }
    }

    if (-not $success) {
        throw "Request failed after $RetryCount attempts due to repeated 429 errors."
    }
}

# Function to retrieve permissions
Function Get-Permissions([Microsoft.SharePoint.Client.SecurableObject]$Object) {
    # Determine the type of the object
    Switch($Object.TypedObject.ToString()) {
        "Microsoft.SharePoint.Client.ListItem" {
            if ($Object.Folder -ne $null) {
                $ObjectType = "Folder"
                Write-Host -f Cyan "Processing permissions for Folder at $($Object.Folder.ServerRelativeUrl)"
                $Object.ParentList.Retrieve("DefaultDisplayFormUrl")
                Execute-QueryWithRetry { $Ctx.ExecuteQuery() }
                $DefaultDisplayFormUrl = $Object.ParentList.DefaultDisplayFormUrl
                $ObjectURL = $("{0}{1}?ID={2}" -f $Ctx.Web.Url.Replace($Ctx.Web.ServerRelativeUrl, ''), $DefaultDisplayFormUrl, $Object.ID)
            } else {
                Write-Host -f Gray "Skipping non-folder item..."
                return  # Skip non-folder items
            }
        }
        Default {
            Write-Host -f Gray "Skipping non-folder object..."
            return  # Skip lists, libraries, and other non-folder objects
        }
    }

    # Get permissions assigned to the object
    $Ctx.Load($Object.RoleAssignments)
    Write-Host -f Yellow "Loading RoleAssignments for $ObjectType at $ObjectURL..."
    Execute-QueryWithRetry { $Ctx.ExecuteQuery() }

    $PermissionsWritten = $false
    foreach ($RoleAssignment in $Object.RoleAssignments) {
        $Ctx.Load($RoleAssignment.Member)
        Write-Host -f Yellow "Loading RoleAssignment Member data..."
        Execute-QueryWithRetry { $Ctx.ExecuteQuery() }

        # Check direct permissions
        if ($RoleAssignment.Member.PrincipalType -eq "User") {
            if ($RoleAssignment.Member.LoginName -eq $SearchUser.LoginName) {
                Write-Host -f Cyan "Found the User under direct permissions of the $($ObjectType) at $($ObjectURL)"
                $UserPermissions = @()
                $Ctx.Load($RoleAssignment.RoleDefinitionBindings)
                Write-Host -f Yellow "Loading RoleDefinitionBindings..."
                Execute-QueryWithRetry { $Ctx.ExecuteQuery() }
                foreach ($RoleDefinition in $RoleAssignment.RoleDefinitionBindings) {
                    $UserPermissions += $RoleDefinition.Name + ";"
                }
                # Send the Data to Report file
                "$($ObjectURL), $($ObjectType), $($Object.Title), Direct Permission, $($UserPermissions)" | Out-File $ReportFile -Append
                $PermissionsWritten = $true
            }
        } elseif ($RoleAssignment.Member.PrincipalType -eq "SharePointGroup") {
            Write-Host -f Cyan "Processing SharePoint Group permissions..."
            $Group = $Web.SiteGroups.GetByName($RoleAssignment.Member.LoginName)
            $GroupUsers = $Group.Users
            $Ctx.Load($GroupUsers)
            Execute-QueryWithRetry { $Ctx.ExecuteQuery() }

            foreach ($User in $GroupUsers) {
                if ($User.LoginName -eq $SearchUser.LoginName) {
                    Write-Host -f Cyan "Found the User under Member of the Group '$($RoleAssignment.Member.LoginName)' on $($ObjectType) at $($ObjectURL)"
                    $GroupPermissions = @()
                    $Ctx.Load($RoleAssignment.RoleDefinitionBindings)
                    Execute-QueryWithRetry { $Ctx.ExecuteQuery() }
                    foreach ($RoleDefinition in $RoleAssignment.RoleDefinitionBindings) {
                        $GroupPermissions += $RoleDefinition.Name + ";"
                    }
                    "$($ObjectURL), $($ObjectType), $($Object.Title), Member of '$($RoleAssignment.Member.LoginName)' Group, $($GroupPermissions)" | Out-File $ReportFile -Append
                    $PermissionsWritten = $true
                }
            }
        }
    }

    if ($PermissionsWritten) {
        Write-Host -f Green "Permissions for $ObjectType at $ObjectURL written to $ReportFile."
    } else {
        Write-Host -f Gray "No permissions to write for $ObjectType at $ObjectURL."
    }
}

# Function to Check Permissions of All List Items of a given List
Function Check-SPOListItemsPermission([Microsoft.SharePoint.Client.List]$List) {
    Write-host -f Yellow "Searching in Folders of the List '$($List.Title)'..."

    $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
    $Query.ViewXml = "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value></Eq></Where><OrderBy><FieldRef Name='ID' Ascending='TRUE'/></OrderBy></Query><RowLimit Paged='TRUE'>$BatchSize</RowLimit></View>"
    $Counter = 0

    # Batch process list items - to mitigate list threshold issue on larger lists
    Do {
        # Get items from the list in Batch
        Write-Host -f Yellow "Loading a batch of folders from the list..."
        $ListItems = $List.GetItems($Query)
        $Ctx.Load($ListItems)
        Execute-QueryWithRetry { $Ctx.ExecuteQuery() }

        $Query.ListItemCollectionPosition = $ListItems.ListItemCollectionPosition
        # Loop through each List item
        foreach ($ListItem in $ListItems) {
            if ($ListItem["FSObjType"] -eq 1) {
                Write-Host -f Cyan "Processing folder item $($ListItem.Id) at $($ListItem["FileRef"])"
                $ListItem.Retrieve("HasUniqueRoleAssignments")
                Execute-QueryWithRetry { $Ctx.ExecuteQuery() }
                if ($ListItem.HasUniqueRoleAssignments -eq $true) {
                    # Call the function to generate Permission report
                    Get-Permissions -Object $ListItem
                } else {
                    Write-Host -f Gray "Folder $($ListItem["FileRef"]) has inherited permissions."
                }
            }
            $Counter++
            Write-Progress -PercentComplete ($Counter / ($List.ItemCount) * 100) -Activity "Processing Folders $Counter of $($List.ItemCount)" -Status "Searching Unique Permissions in Folders of '$($List.Title)'"
        }
    } while ($Query.ListItemCollectionPosition -ne $null)
}

# Function to Check Permissions of all lists from the web
Function Check-SPOListPermission([Microsoft.SharePoint.Client.Web]$Web) {
    # Get All Lists from the web
    $Lists = $Web.Lists
    $Ctx.Load($Lists)
    Execute-QueryWithRetry { $Ctx.ExecuteQuery() }

    # Get all lists from the web 
    foreach ($List in $Lists) {
        # Exclude System Lists
        if ($List.Hidden -eq $False) {
            # Get List Items Permissions
            Check-SPOListItemsPermission $List

            # Get the Lists with Unique permission
            $List.Retrieve("HasUniqueRoleAssignments")
            Execute-QueryWithRetry { $Ctx.ExecuteQuery() }

            if ($List.HasUniqueRoleAssignments -eq $True) {
                # Call the function to check permissions
                Get-Permissions -Object $List
            }
        }
    }
}

# Function to Check Web's Permissions from given URL
Function Check-SPOWebPermission([Microsoft.SharePoint.Client.Web]$Web) {
    # Get all immediate subsites of the site
    $Ctx.Load($web.Webs)
    Execute-QueryWithRetry { $Ctx.ExecuteQuery() }

    # Call the function to Get Lists of the web
    Write-host -f Yellow "Searching in the Web $($Web.URL)..."

    # Check if the Web has unique permissions
    $Web.Retrieve("HasUniqueRoleAssignments")
    Execute-QueryWithRetry { $Ctx.ExecuteQuery() }

    # Get the Web's Permissions
    if ($web.HasUniqueRoleAssignments -eq $true) {
        Get-Permissions -Object $Web
    }

    # Scan Lists with Unique Permissions
    Write-host -f Yellow "Searching in the Lists and Libraries of $($Web.URL)..."
    Check-SPOListPermission($Web)

    # Iterate through each subsite in the current web
    foreach ($Subweb in $web.Webs) {
        # Call the function recursively
        Check-SPOWebPermission $SubWeb
    }
}

# Main script execution
Try {
    # Authenticate to SharePoint Online using PnP PowerShell
    Connect-PnPOnline -Url $SiteURL -Interactive
    $Ctx = Get-PnPContext

    # Get the Web
    $Web = $Ctx.Web
    $Ctx.Load($Web)
    Execute-QueryWithRetry { $Ctx.ExecuteQuery() }

    # Prompt user to enter the account to search for
    $UserAccount = Read-Host "Enter the user account (e.g., user@domain.com) to search permissions for"

    # Get the User object
    $SearchUser = $Web.EnsureUser($UserAccount)
    $Ctx.Load($SearchUser)
    Execute-QueryWithRetry { $Ctx.ExecuteQuery() }

    # Write CSV (comma-separated file) Header
    "URL,Object,Title,PermissionType,Permissions" | Out-File $ReportFile

    Write-host -f Yellow "Searching in the Site Collection Administrators Group..."
    # Check if Site Collection Admin
    if ($SearchUser.IsSiteAdmin -eq $True) {
        Write-host -f Cyan "Found the User under Site Collection Administrators Group!"
        # Send the Data to report file
        "$($Web.URL),Site Collection,$($Web.Title),Site Collection Administrator,Site Collection Administrator" | Out-File $ReportFile -Append
    }

    # Call the function with RootWeb to get site collection permissions
    Check-SPOWebPermission $Web

    Write-host -f Green "User Permission Report Generated Successfully!"
}
Catch {
    write-host -f Red "Error Generating User Permission Report!" $_.Exception.Message
}
