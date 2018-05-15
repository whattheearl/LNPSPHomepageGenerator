#d/l file <---- should put up a template ---> or save one created so can just replace text!
#Get-PnPFile -url /sites/lnp/pages/Home.aspx -AsFile
$cred = Get-PnPStoredCredential -Name LairdNorton -Type PSCredential
Connect-PnPOnline "https://lnco.sharepoint.com/sites/lnp" -Credentials $cred

#check out file
$rel_home_url = "/sites/lnp/pages/Home.aspx"
set-pnpfilecheckedout -Url $rel_home_url

#upload file
Add-PnPFile -Path ".\out\Home.aspx" -Folder "/Pages"

#check in file
set-pnpfilecheckedin -Url $rel_home_url

#ensure correct homepage
Set-PnPHomePage -RootFolderRelativeUrl 'pages/Home.aspx'