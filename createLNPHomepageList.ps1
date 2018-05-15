$outfile = "./out/test.html"
new-item $outfile -force

$cred = Get-PnPStoredCredential -name LairdNorton -type PSCredential

write-host "Getting Properties"
$rootUrl = "https://lnco.sharepoint.com/sites/lnp"
Write-Host "Connecting to $($rootUrl)"
Connect-PnPOnline $rootUrl -Credentials $cred

Write-Host "Getting Subsites"
$items = Get-PnPListItem -List "Lists/PropertiesList" -Fields "PicsLibrary", "DocsLibrary", "Title", "chRealEstatePortfolio", "chAcquisitionStage"

class PropertiesListItem 
{
    $Data
    PropertiesListItem (){}
    [string] Portfolio() { return [string]$this.Data["chRealEstatePortfolio"].label }
    [string] PicsLibrary() { return [string]$this.Data["PicsLibrary"].url; }
    [string] DocsLibrary() { return [string]$this.Data["DocsLibrary"].url; }
    [string] Title() { return [string]$this.Data["Title"]; }
    [string] AquisitionStage() { return [string]$this.Data["chAcquisitionStage"].label }
}



$properties = @()
foreach($item in $items) 
{
    $object = new-object PropertiesListItem
    $object.Data = $item
    $properties += $object
}

$spectrumBlank = @()
$spectrumOwned = @()
$lanogaOwned = @()
$lanogaSold = @()
$newOpportunity = @()
$unicoOwned = @()
foreach($prop in $properties) {
    switch($prop.Portfolio())
    {
        "Lanoga" { 
            if( $prop.AquisitionStage() -eq "Owned") { $lanogaOwned += $prop }
            elseif ( $prop.AquisitionStage() -eq "Sold") { $lanogaSold += $prop}
        }
        "Unico JV" { 
            if( $prop.AquisitionStage() -eq "Owned") { $unicoOwned += $prop }
            write-host "unico" + $prop.Title() $prop.Portfolio()
        }
        "Spectrum" {
            if( $prop.AquisitionStage() -eq "Owned") { $spectrumOwned += $prop; write-host $spectrumOwned; write-host $prop; }
            elseif( $prop.AquisitionStage() -eq "") { $spectrumBlank += $prop}
            else { write-host $prop.AquisitionStage }
            write-host $prop.Title() $prop.Portfolio()
        }
        "Direct Investments" { $newOpportunity += $prop }
    } 
}
#start of table
write-host "sold " + $lanogaSold.length
write-host "owned " + $lanogaOwned.length

function CSS() {
    return '<style>
    .col {
        float:left;
        width: 30%;
        margin-right: 3%;
        min-width: 250px;
    }
    
    table {
        margin-bottom: 10px;
    }

    .title {
        line-height: 12px;
    }
    
    h2.table-title {
        border-collapse: collapse;
        color: rgb(51, 153, 0);
        overflow-wrap: break-word;
        text-decoration: none solid rgb(51, 153, 0);
        word-wrap: break-word;
        column-rule-color: rgb(51, 153, 0);
        caret-color: rgb(51, 153, 0);
        border: 0px none rgb(51, 153, 0);
        font: normal normal 200 normal 25.2px / 35.28px "Segoe UI Light", "Segoe UI", Segoe, Tahoma, Helvetica, Arial, sans-serif;
        outline: rgb(51, 153, 0) none 0px;
        margin: 0;
        margin-bottom: 5px;
    }
    
    td:first-child { 
        min-width: 150px;
    }

    .padding{
        
    }
    th, td {
        padding: 5px 5px 5px 0;
        font-size: 14px;
        font-family: "Segoe UI";        
    }

    th {
        color: rgb(102, 102, 102);
        cursor: default;
        overflow-wrap: break-word;
        text-align: left;
        width: 100/3%;
        text-decoration: none solid rgb(102, 102, 102);
        white-space: nowrap;
        word-wrap: break-word;
        column-rule-color: rgb(102, 102, 102);
        caret-color: rgb(102, 102, 102);
        border: 0px none rgb(102, 102, 102);
        outline: rgb(102, 102, 102) none 0px;
    }/*#A_1*/
    
    tr {
        width: 100%;
    }
    
    td {
        color: rgb(102, 102, 102);
        overflow-wrap: break-word;
        text-decoration: none solid rgb(102, 102, 102);
        text-align: left;
        word-wrap: break-word;
        column-rule-color: rgb(102, 102, 102);
        caret-color: rgb(102, 102, 102);
        outline: rgb(102, 102, 102) none 0px;
    }
    
    td a {
        color: rgb(51, 153, 0);
        overflow-wrap: break-word;
        text-decoration: none solid rgb(51, 153, 0);
        word-wrap: break-word;
        column-rule-color: rgb(51, 153, 0);
        caret-color: rgb(51, 153, 0);
        border: 0px none rgb(51, 153, 0);
        outline: rgb(51, 153, 0) none 0px;
    }
    </style>'
}

function TableHead( $tableName ) {
    return '<h2 class="table-title">' + $tableName + '</h2>
    <table>
    <thead>
      <tr>
        <th>Property Name</th>
        <th>DocsLibrary</th>
        <th>PicsLibrary</th>
      </tr>
    </thead>
    <tbody>'
}

function TableBody($pListItems) 
{
    [string]$ans = ""
    for($i = 0; $i -lt $pListItems.Count; $i++)
    {
        $ans += '<tr>
            <td class="title">'+$pListItems[$i].Title()+'</td>
            <td><a href="'+$pListItems[$i].DocsLibrary()+'">Documents</a></td>
            <td><a href="'+$pListItems[$i].PicsLibrary()+'">Pictures</a></td>
        </tr>'
    }
    return $ans
}

function TableFoot() {
    return '</tbody>
    </table>'
}

function WriteLanogaOwned($lanogaOwned){
    '<section class="col">' >> $outfile
    TableHead('Lanoga Properties Owned') >> $outfile
    TableBody($lanogaOwned) >> $outfile
    TableFoot >> $outfile
    '</section>' >> $outfile
}

function WriteLanogaSold($lanogaSold){
    '<section class="col">' >> $outfile
    TableHead('Lanoga Properties Sold') >> $outfile
    TableBody($lanogaSold) >> $outfile
    TableFoot >> $outfile
    '</section>' >> $outfile
}

function WriteRightCol($arrays){
    write-host "inside writeright"
    '<section class="col">' >> $outfile
    TableHead('Direct Investments') >> $outfile
    TableBody($arrays[0]) >> $outfile
    TableFoot >> $outfile
    TableHead('Unico Properties Owned') >> $outfile
    TableBody($arrays[1]) >> $outfile
    TableFoot >> $outfile
    TableHead('Spectrum Owned') >> $outfile
    TableBody($arrays[2]) >> $outfile
    TableFoot >> $outfile
    TableHead('Spectrum') >> $outfile
    TableBody($arrays[3]) >> $outfile
    TableFoot >> $outfile
    '</section>' >> $outfile
}

write-host "Writing CSS"
"<span id='ms-rterangeselectionplaceholder-start'></span><span id='ms-rterangeselectionplaceholder-end'></span>" >> $outfile
CSS >> $outfile
write-host "Writing Lanoga Owned"
WriteLanogaOwned($lanogaOwned)
write-host "Writing Lanoga Sold"
WriteLanogaSold($lanogaSold)
write-host "Writing newopp unicoowned spectrum"
WriteRightCol(@($newOpportunity, $unicoOwned, $spectrumOwned, $spectrumBlank))
write-host "converting charaters"
py ./EscapeHtml.py

write-host "uploading files"
./Upload-File