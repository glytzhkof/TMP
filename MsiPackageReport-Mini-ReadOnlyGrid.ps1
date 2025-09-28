#Add-Type -AssemblyName WindowsBase

# Clean the Powershell console window
Clear-Host

# Windows Installer COM object (MSI is old)
$installer = New-Object -ComObject WindowsInstaller.Installer

$msiUILevelNone = 2 # Show no GUI for activated MSI Session objects
$p = 1

#$ErrorActionPreference = 'SilentlyContinue' # Just continue with next package on error

# $userInput = Read-Host "Please enter your name"
# Write-Host "Hello, $userInput!"


<# $result = [System.Windows.Forms.MessageBox]::Show(
    "Do you want to save your changes?",
    "Save Confirmation",
    [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
    [System.Windows.Forms.MessageBoxIcon]::Question,
    [System.Windows.Forms.MessageBoxDefaultButton]::Button1,
    [System.Windows.Forms.MessageBoxOptions]::DefaultDesktopOnly
)

switch ($result) {
    'Yes' {
        Write-Host "User chose Yes."
    }
    'No' {
        Write-Host "User chose No."
    }
    'Cancel' {
        Write-Host "User chose Cancel."
    }
} #>

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

# Get all installed MSI packages and prepare to initiate session object in no-GUI mode
$products = $installer.ProductsEx("", "", 7)
$totalpackages = $products.Count()
$installer.UILevel = $msiUILevelNone #[Type]::GetType("Microsoft.Deployment.WindowsInstaller.InstallerUILevel").GetField("None").GetValue($null)

# Empty array to hold product codes that share the same upgrade code (related products)
$relatedproductcodes = @()

# A class to hold MSI package information
class MsiPackage {
    [string]$Counter; [string]$ProductName; [string]$Version; [string]$PackageCode; [string]$ProductCode
    [string]$UpgradeCode ; [string]$RelatedProductCodes; [string]$Assignment; [string]$Lcid
} # Assignment = Installation Scope translated to text (user or machine)

# We need a collection for the package information
$MsiPackages = New-Object System.Collections.Generic.List[MsiPackage]

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
        $allupgrades = $relatedproductcodes -join "`r`n"
    }
    
    # Add to MSI package list for Powershell grid exports and standard HTML output 
    $MsiPackages.Add([MsiPackage]@{Counter=$p;ProductName=$productname;Version=$versionstring;PackageCode=$packagecode;ProductCode=$productcode;UpgradeCode=$upgradecode;RelatedProductCodes=$allupgrades;Assignment=$assignment;Lcid=$lcid})

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

Write-Host "Generating output grid..."

# Show read-only grid view - just dump the whole package list to the Out-Gridview Cmdlet
$MsiPackages | Out-GridView -PassThru # -PassThru to avoid script window closing if run from right click
#$MsiPackages | Out-GridView

Write-Host "Execution complete."

#Read-Host -Prompt "Press Enter to exit"

<# if ($Error)
{
    Pause
} #>

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