# Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypas

Add-Type -assembly System.Windows.Forms

# Define the global variable for the abort function
$global:AbortIsoCreation = $false 

# Define the error handling function that displays to GUI
function Show-ErrorDialog {
    param(
        [string]$ErrorMessage
    )
    [System.Windows.Forms.MessageBox]::Show($Errormessage, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}

# Define the message function that displays to GUI
function Show-Infomessage {
    param(
        [string]$Message
    )
    [System.Windows.Forms.MessageBox]::Show($Message, "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}

# Abort function
function abort-IsoCreation{
    $global:AbortIsoCreation = $true
}

# Define the ISO function
function New-IsoFileWithGUI {
    [CmdletBinding()]
    Param(
        [string]$SourceDir,
        [string]$DestinationDir,
        [string]$MediaType
    )

    # Validate source directory
    if (-not (Test-Path -Path $SourceDir -PathType Container)) {
        Show-ErrorDialog "Source directory '$SourceDir' does not exist or is not accessible."
        # SEND THIS ERROR TO GUI
        return
    }

    # Validate destination directory
    if (-not (Test-Path -Path $DestinationDir -PathType Container)) {
        Show-ErrorDialog "Destination directory '$DestinationDir' does not exist or is not accessible."
        # SEND THIS ERROR TO GUI
        return
    }

    # Function definition
    function New-IsoFile 
    {   
        [CmdletBinding(DefaultParameterSetName='Source')]
        Param( 
            [parameter(Position=1,Mandatory=$true,ValueFromPipeline=$true, ParameterSetName='Source')]
            $Source,  
            [parameter(Position=2)]
            [string]$Path = "$DestinationDir\$((Get-Date).ToString('yyyyMMdd-HHmmss.ffff')).iso",  
            [ValidateScript({Test-Path -LiteralPath $_ -PathType Leaf})]
            [string]$BootFile = $null, 
            [ValidateSet('CDR','CDRW','DVDRAM','DVDPLUSR','DVDPLUSRW','DVDPLUSR_DUALLAYER','DVDDASHR','DVDDASHRW','DVDDASHR_DUALLAYER','DISK','DVDPLUSRW_DUALLAYER','BDR','BDRE')]
            [string] $Media = 'DVDPLUSRW_DUALLAYER', 
            [string]$Title = (Get-Date).ToString("yyyyMMdd-HHmmss.ffff"),  
            [switch]$Force, 
            [parameter(ParameterSetName='Clipboard')]
            [switch]$FromClipboard 
        ) 

        Begin {  
            ($cp = new-object System.CodeDom.Compiler.CompilerParameters).CompilerOptions = '/unsafe' 
            if (!('ISOFile' -as [type])) {  
                Add-Type -CompilerParameters $cp -TypeDefinition @'
                public class ISOFile  
                { 
                    public unsafe static void Create(string Path, object Stream, int BlockSize, int TotalBlocks)  
                    {  
                        int bytes = 0;  
                        byte[] buf = new byte[BlockSize];  
                        var ptr = (System.IntPtr)(&bytes);  
                        var o = System.IO.File.OpenWrite(Path);  
                        var i = Stream as System.Runtime.InteropServices.ComTypes.IStream;  

                        if (o != null) { 
                            while (TotalBlocks-- > 0) {  
                                i.Read(buf, BlockSize, ptr); o.Write(buf, 0, bytes);  
                            }  
                            o.Flush(); o.Close();  
                        } 
                    } 
                }  
'@  
            } 

            if ($BootFile) { 
                if('BDR','BDRE' -contains $Media) { Write-Warning "Bootable image doesn't seem to work with media type $Media" } 
                ($Stream = New-Object -ComObject ADODB.Stream -Property @{Type=1}).Open()  # adFileTypeBinary 
                $Stream.LoadFromFile((Get-Item -LiteralPath $BootFile).Fullname) 
                ($Boot = New-Object -ComObject IMAPI2FS.BootOptions).AssignBootImage($Stream) 
            } 

            $MediaType = @('UNKNOWN','CDROM','CDR','CDRW','DVDROM','DVDRAM','DVDPLUSR','DVDPLUSRW','DVDPLUSR_DUALLAYER','DVDDASHR','DVDDASHRW','DVDDASHR_DUALLAYER','DISK','DVDPLUSRW_DUALLAYER','HDDVDROM','HDDVDR','HDDVDRAM','BDROM','BDR','BDRE') 

            # Show-InfoMessage "Selected media type is $Media with value $($MediaType.IndexOf($Media))"
            ($Image = New-Object -com IMAPI2FS.MsftFileSystemImage -Property @{VolumeName=$Title}).ChooseImageDefaultsForMediaType($MediaType.IndexOf($Media)) 

            if (!($Target = New-Item -Path $Path -ItemType File -Force:$Force -ErrorAction SilentlyContinue)) { 
                Show-ErrorDialog "Cannot create file $Path. Use -Force parameter to overwrite if the target file already exists."; 
                break 
            } 
        }  

        Process { 
            if($global:AbortIsoCreation){
                Show-InfoMessage "ISO creation process aborted."
                return
            }

            if($FromClipboard) { 
                if($PSVersionTable.PSVersion.Major -lt 5) { Show-ErrorDialog'The -FromClipboard parameter is only supported on PowerShell v5 or higher'; break } 
                $Source = Get-Clipboard -Format FileDropList 
            } 

            foreach($item in $Source) { 
                if($item -isnot [System.IO.FileInfo] -and $item -isnot [System.IO.DirectoryInfo]) { 
                    $item = Get-Item -LiteralPath $item
                } 

                if($item) { 
                    Write-Verbose -Message "Adding item to the target image: $($item.FullName)"
                    try { $Image.Root.AddTree($item.FullName, $true) } catch { Show-ErrorDialog ($_.Exception.Message.Trim() + ' Try a different media type.') } 
                } 
            } 
        } 

        End {  
            if ($Boot) { $Image.BootImageOptions=$Boot }  
            $Result = $Image.CreateResultImage()  
            [ISOFile]::Create($Target.FullName,$Result.ImageStream,$Result.BlockSize,$Result.TotalBlocks) 
            Show-InfoMessage "Target image ($($Target.FullName)) has been created"
            $Target
        } 
    } 

    $source_dir = $SourceDir
    New-ISOFile -Source $source_dir -Path $DestinationDir -Media $MediaType
}

# Create the form

$form = New-Object System.Windows.Forms.Form
$form.Text = 'GUI for MakeISO'
#set window pop-up size
$form.Width = 600
$form.Height = 400
# Autosize automatically expands the window if the elements exceed the boundary
$form.Autosize = $true

# Define a global variable to store the selected source items
$global:SelectedSourceItems = @()

# Select Source Directory Label
$LabelSource = New-Object System.Windows.Forms.Label
$LabelSource.Text = "Source Directory:"
$LabelSource.Left = 10 
$LabelSource.Top = 20
$LabelSource.Width = 130
$form.Controls.Add($LabelSource)

# Select Destination Directory Label
$LabelDestination = New-Object System.Windows.Forms.Label
$LabelDestination.Text = "Destination Directory:"
$LabelDestination.Left = 10 
$LabelDestination.Top = 50
$LabelDestination.Width = 130
$form.Controls.Add($LabelDestination)

# Create textboxes
$txtSource = New-Object System.Windows.Forms.TextBox
$txtSource.Left = 150
$txtSource.Top = 20
$txtSource.Width = 250
$form.Controls.Add($txtSource)

$txtDestination = New-Object System.Windows.Forms.TextBox
$txtDestination.Left = 150
$txtDestination.Top = 50
$txtDestination.Width = 250
$form.Controls.Add($txtDestination)


# Create browse buttons
$btnBrowseSource = New-Object System.Windows.Forms.Button
$btnBrowseSource.Left = 410
$btnBrowseSource.Top = 20
$btnBrowseSource.Text = "Browse..."
$btnBrowseSource.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select Source Directory"
    $folderBrowser.SelectedPath = $txtSource.Text
    $result = $folderBrowser.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtSource.Text = $folderBrowser.SelectedPath 
    }
}) 
$form.Controls.Add($BtnBrowseSource)

$btnBrowseDestination = New-Object System.Windows.Forms.Button
$btnBrowseDestination.Left = 410
$btnBrowseDestination.Top = 50
$btnBrowseDestination.Text = "Browse..."
$btnBrowseDestination.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select Destination Directory"
    $folderBrowser.SelectedPath = $txtDestination.Text
    $result = $folderBrowser.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtDestination.Text = $folderBrowser.SelectedPath
    }
}) 
$form.Controls.Add($btnBrowseDestination)

# Create media type dropdown
$cmbMediaType = New-Object System.Windows.Forms.ComboBox
$cmbMediaType.Left = 150
$cmbMediaType.Top = 80
$cmbMediaType.Width = 250
$cmbMediaType.Items.AddRange(@("CDR","CDRW","DVDRAM","DVDPLUSR","DVDPLUSRW","DVDPLUSR_DUALLAYER","DVDDASHR","DVDDASHRW","DVDDASHR_DUALLAYER","DISK","DVDPLUSRW_DUALLAYER","BDR","BDRE"))
$form.Controls.Add($cmbMediaType)

# Select Destination Directory
$LabelMedia = New-Object System.Windows.Forms.Label
$LabelMedia.Text = "Media Type:"
$LabelMedia.Left = 10 
$LabelMedia.Top = 80
$form.Controls.Add($LabelMedia)

# Create button to start ISO creation
$btnCreateISO = New-Object System.Windows.Forms.Button
$btnCreateISO.Left = 150
$btnCreateISO.Top = 120
$btnCreateISO.Width = 100
$btnCreateISO.Text = "Create ISO"
$btnCreateISO.Add_Click({
    # Reset abort flag
    $global:AbortIsoCreation = $false 
    $sourceDir = $txtSource.Text
    $destinationDir = $txtDestination.Text
    $mediaType = $cmbMediaType.SelectedItem

    New-IsoFileWithGUI -SourceDir $sourceDir -DestinationDir $destinationDir -MediaType $mediaType
})
$form.Controls.Add($btnCreateISO)

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Style = 'Continuous'
$ProgressBar.Left = 150
$ProgressBar.Top = 160
$progressBar.Width = 250
$form.Controls.Add($progressBar)

# Select Progress bar
$LabelProgress = New-Object System.Windows.Forms.Label
$LabelProgress.Text = "Progress:"
$LabelProgress.Left = 10 
$LabelProgress.Top = 160
$form.Controls.Add($LabelProgress)

# Create abort button
$buttonAbort = New-Object system.Windows.Forms.Button
$buttonAbort.Left = 270
$buttonAbort.Top = 120
$buttonAbort.width = 100
$buttonabort.Text = "Abort"
$buttonAbort.Add_Click({
    Abort-IsoCreation
    
})
$form.Controls.Add($buttonAbort)


# Show the form
$form.ShowDialog() | Out-Null


