#Add-Type -AssemblyName WindowsBase

# Clean the Powershell console window
Clear-Host

# Windows Installer COM object (MSI is old)
$installer = New-Object -ComObject WindowsInstaller.Installer

$msiUILevelNone = 2 # Show no GUI for activated MSI Session objects
$p = 1

#$ErrorActionPreference = 'SilentlyContinue' # Just continue with next package on error

# Show some warnings and allow cancelling script
Add-Type -AssemblyName System.Windows.Forms
$msgBody = "This export may take quite some time to complete.`n`nPlease click OK and wait for the results to appear in your browser, or click Cancel to exit without running the script."
$msgTitle = "MSI Info Export Starting"
$msgButton = "OKCancel"
$msgIcon = "Question"
$result = [System.Windows.Forms.MessageBox]::Show($msgBody, $msgTitle, $msgButton, $msgIcon)
Write-Host "The user selected: $Result"
if ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
    Write-Host "User clicked Cancel. Ending script."
    exit
} 
Write-Host "Continuing..."

# Construct the html header for output file (using here-string)
$htmloutput = @"
<!DOCTYPE html>
<html lang='en'><head><title>MSI Package Estate Information:</title><meta charset='utf-8'>
<script>function init() { try { document.querySelectorAll('td').forEach(link => { link.addEventListener('mouseenter', function (event) {var range = document.createRange(); range.selectNodeContents(this); var sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(range);});});} catch (error) { console.log(error); }}
function filterTable(filter) { var row; var rows = document.querySelectorAll('table tbody tr'); var rowcount = rows.length; var hiddenrows = 0; for (row = 0; row < rowcount; row++) { if (rows[row].textContent.toUpperCase().indexOf(filter.toUpperCase()) > -1) { rows[row].style.display = '';} else { rows[row].style.display = 'none'; hiddenrows++;}}}
function reset() {document.getElementById('search-box').value = '';filterTable('');}</script>
<style>body {font: 12px Calibri;}a {color: lightgrey;} a:hover {background-color: black;}
table, td {border: 1px solid black;border-collapse: collapse;padding: 0.3em;vertical-align: text-top;border-top: none;}
table>*>tr>td:nth-child(2) { max-width: 300px;}
th {font: bold 18px Calibri;background-color: purple;text-align: left;color: white;}
table th {position: sticky;top: -1px;}</style>
</head><body onload='init()'>
<h1>MSI Package Report</h1><input id='search-box' type='text' onemptied='reset()' autocomplete='off' oninput='filterTable(this.value)' title='Filter table by keyword search' placeholder='Filter by...'>
<button onclick='reset()'>x</button><h3>Use your browser's zoom setting to make text more readable.</h3>
<table><thead><tr>
<th>#</th><th>Product Name</th><th>Version</th><th>Package Code</th><th>Product Code</th><th>Upgrade Code</th><th  title='Product codes that share the same upgrade code.'>Related Product Codes</th><th>Scope</th><th><a href='https://msdn.microsoft.com/en-us/library/ms912047(v=winembedded.10).aspx' target='_blank'>LCID</a></th>
</tr></thead><tbody>`r`n
"@

# Get all installed MSI packages and prepare to initiate session object in no-GUI mode
$products = $installer.ProductsEx("", "", 7)
$totalpackages = $products.Count()
$installer.UILevel = $msiUILevelNone #[Type]::GetType("Microsoft.Deployment.WindowsInstaller.InstallerUILevel").GetField("None").GetValue($null)

# Empty array to hold product codes that share the same upgrade code (related products)
$relatedproductcodes = @()

# Now process each MSI package in sequence
foreach ($product in $products) {

    $productcode = $product.ProductCode() # Crucial: must add () at end even if it is a property in the object model
    $productname = $product.InstallProperty('ProductName')
    $versionstring = $product.InstallProperty('VersionString')
    $packagecode = $product.InstallProperty('PackageCode')
    $scope = $product.InstallProperty("AssignmentType")
    $lcid = $product.InstallProperty("Language")
    $upgradecode = "" # Will be retrieved later

    switch ($scope) {
        0 { $assignment = "User" }
        1 { $assignment = "Machine" }
        default { $assignment = "Unknown" }
    }

    try {
        # Get upgrade code via MSI session object (reads cached MSI database with applied transforms - apparently)
        $session = $installer.OpenProduct($productcode)

        # So far so good, we have our session object, but upgrade code can be missing 
        $upgradecode = $session.ProductProperty("UpgradeCode")

        # Don't pass empty string to RelatedProducts, a runtime error will result
        if ($upgradecode -ne "") {
            # RelatedProducts lists products that share the same upgrade code (they are related)
            $upgrades = $installer.RelatedProducts($upgradecode)
            foreach ($u in $upgrades) {
                $relatedproductcodes += $u
            }
        }
    }
    catch {
       # Our whole session object failed to instantiate, report error in export
       $upgradecode = "Error Accessing Data: $($_.Exception.Source), 0x$([Convert]::ToString($_.Exception.HResult,16))"
    }
    finally {
        # Crucial: Always release the session object in order to be able to continue with
        #          the next package regardless if there was an error or not (hence finally)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($session) | Out-Null
    }

    # Create html element listing all related product codes (if more than one)
    if ($relatedproductcodes.Count -gt 0) {
        $allupgrades = $relatedproductcodes -join "<br>"
    }
    
    # The MSI package details we want to output for this product (for custom HTML report)
    $htmloutput += "<tr><td>$p</td><td>$productname</td><td>$versionstring</td><td>$packagecode</td><td>$productcode</td><td>$upgradecode</td><td>$allupgrades</td><td>$assignment</td><td>$lcid</td></tr>`r`n"
 
    # Clean up things for next package
    $relatedproductcodes = @()
    $upgradecode = ""
    $allupgrades = ""

    $p++

    # Show a progress bar for the package retrieval process
    $progress = [math]::Floor(($p / $totalpackages) * 100)
    Write-Progress -Activity "Package retrieval" "$progress % Complete:" -percentComplete $progress;
}

# Update display and remove progress bar
Write-Progress -Activity "Package retrieval" -Completed

# Release Windows Installer COM object as early as possible
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($installer) | Out-Null

# Finalize the custom HTML output file content
$htmloutput += "</tbody></table></body></html>"

# Build output filename with computer name, date and time embedded in filename for custom HTML export
$filename = "MsiInfo_$($env:COMPUTERNAME)_$((Get-Date).Day).$((Get-Date).Month)(month).$((Get-Date).Year)_$((Get-Date).Hour)-$((Get-Date).Minute)-$((Get-Date).Second).html"

# Create output file for custom HTML export
Write-Host "Generating output file..."
$outputpath = Join-Path -Path $PSScriptRoot -ChildPath $filename
$Utf8BomEncoding = New-Object System.Text.UTF8Encoding(1) # Using Utf8 with BOM - for now...
[System.IO.File]::WriteAllLines($outputpath, $htmloutput, $Utf8BomEncoding)

# Open the custom, exported HTML output file in default browser
Start-Process $outputpath

Write-Host "Execution complete."

<# clean {
   Write-Host "Executing cleanup block..."
   #Release Windows Installer COM object if script crashed
   [System.Runtime.InteropServices.Marshal]::ReleaseComObject($installer) | Out-Null
#  [System.GC]::Collect()
} #>

#$products = $installer.ProductsEx("", "", 7)
#$products.Count
#write-host $products.Count()
    #Write-Progress -Activity "Package retrieval" "$p % Complete:" -percentComplete $p;
#$products = $inst.Products
#$inst | Get-Member -MemberType Property
#$comObject | Get-Member -MemberType Property
#[System.Windows.MessageBox]::Show([int]$inst.Products.Count).ToString()
#| write-host
# $tmp = $installer.Products.Count.ToString()
# write-host $tmp
#     Dim inst As WindowsInstaller.Installer
# Set inst = CreateObject("WindowsInstaller.Installer")
# MsgBox CStr(inst.Products.Count)
