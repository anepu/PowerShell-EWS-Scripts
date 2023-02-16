#
# Stamp-Folder-with-Personal-Tag.ps1
#
# By Andrei Epure. Use at your own risk.  No warranties are given.
# The script will impersonate a specific mailbox, will target a User Created folder and will stamp a Personal Retention tag on that folder and all of its subfolders
#
# DISCLAIMER:
# THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
# MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
# A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
# MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
# BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
# SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
# OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.

#You need to go through this article to set up an EWS application by using OAuth https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth
#Also the admin that will perform the operation needs to have impersonation role assigned

Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"

# Request parameters that need to be filled in with your own identifiers
$clientId = "00000000-0000-0000-0000-000000000000"
$clientSecret = "your_client_secret"
$adminusername = "admin@contoso.onmicrosoft.com"
$impersonatedAccount = "user@contoso.onmicrosoft.com"
$folderName = "Test"
$PolicyTagRetentionID = "513ea126-7f62-4fbf-a66d-72231ab0f2f3" #Get-RetentionPolicyTag -Identity "name_of_the_tag" | FL Identity, RetentionId
#$PolicyTagRetentionID = $null
# End of parameters

$resource = "https://outlook.office365.com/"
$secureString = Read-Host -Prompt "Enter your Microsoft 365 Admin password" -AsSecureString
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString)
$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

# Token endpoint URL
$tokenEndpoint = "https://login.windows.net/common/oauth2/token"

# Request body
$body = @{
    grant_type = "password"
    client_id = $clientId
    client_secret = $clientSecret
    resource = $resource
    username = $adminusername
    password = $password
}

# Get OAuth2 access token
$response = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $body
$accessToken = $response.access_token

# Create an instance of the ExchangeService object
$exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::ExchangeOnline)

# Set the URL for the Exchange Web Services (EWS) endpoint
$exchangeService.Url = New-Object Uri("https://outlook.office365.com/EWS/Exchange.asmx")

# Set the OAuth2 access token for authentication
$exchangeService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($accessToken)

# Set the impersonated account to retrieve the folder and its subfolders from
$exchangeService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $impersonatedAccount)

# Set the folder name you wish to change the tag for
$folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(500)
$folderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
$searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $folderName)
$findFoldersResults = $exchangeService.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $searchFilter, $folderView)

# Find all subfolders for the primary folder above
$resultFolders = @()
foreach ($folder in $findFoldersResults.Folders) {
    $resultFolders += $folder
    $subfolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
    $subfolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
    $subfolderResults = $exchangeService.FindFolders($folder.Id, $subfolderView)
    foreach ($subfolder in $subfolderResults.Folders) {
        $resultFolders += $subfolder
    }
}

# Stamp the tag on each of the found subfolder
ForEach ($folderfound in $resultFolders)
{
        #PR_POLICY_TAG 0x3018
        $PolicyTag = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3018,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);

        #PR_RETENTION_FLAGS 0x301D    
        $RetentionFlags = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301D,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
        
        #PR_RETENTION_PERIOD 0x301A
        $RetentionPeriod = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301A,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);

        #Bind to the folder found
        $ofolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeservice,$folderfound.Id)
       
        #Same as the value in the PR_RETENTION_FLAGS property
        $ofolder.SetExtendedProperty($RetentionFlags, 144)

        #Same as the value in the PR_RETENTION_PERIOD property
        $ofolder.SetExtendedProperty($RetentionPeriod, 1095)

        #Change the GUID based on your policy tag
        #You need to use a RetentionTag that is available in the currently assigned RetentionPolicy
        $PolicyTagGUID = new-Object Guid($PolicyTagRetentionID);

        $ofolder.SetExtendedProperty($PolicyTag, $PolicyTagGUID.ToByteArray())
        Write-host "Retention policy stamped on" $folderfound.DisplayName -ForegroundColor Green

#Apply the new retention tag
        $ofolder.Update()

}

# Display the folder and all of its subfolders (recursive)
$resultFolders | Select-Object -Property DisplayName, TotalCount, ChildFolderCount

###Verify the Personal tag was applied to the main folder and all of its subfolders (you need to connect to ExO for this)
#Get-MailboxFolderStatistics $impersonatedAccount | where {$_.ArchivePolicy -match "name of policy"}| fl Name,ContainerClass,ArchivePolicy,FolderType
