# Input Parameters  
$ClientId = "f8d6c806-f3e5-49ab-839e-688819832ca4"  
$ClientSecret ="Db9E-Xpg09POSCj4~qW38q.wp-1NHfb-5X"
$TenantId = "b6c37983-27f4-4c9c-9ab2-d2ae66994bc7"  

# Create a hashtable for the body, the data needed for the token request
# The variables used are explained above
$Body = @{
    'tenant' = $TenantId
    'client_id' = $ClientId
    'scope' = 'https://graph.microsoft.com/.default'
    'client_secret' = $ClientSecret
    'grant_type' = 'client_credentials'
}

# Assemble a hashtable for splatting parameters, for readability
# The tenant id is used in the uri of the request as well as the body
$Params = @{
    'Uri' = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    'Method' = 'Post'
    'Body' = $Body
    'ContentType' = 'application/x-www-form-urlencoded'
}

$AuthResponse = Invoke-RestMethod @Params


$Headers = @{
    'Authorization' = "Bearer $($AuthResponse.access_token)"
}

$Result = Invoke-RestMethod -Uri 'https://graph.microsoft.com/v1.0/users' -Headers $Headers
$Result