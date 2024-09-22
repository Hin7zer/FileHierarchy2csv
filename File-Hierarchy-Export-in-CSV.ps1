####################################
#   File Hierarchy  Export as CSV  #
#   Author  Benjamin Hinz          #
#   Version:       1.0             #
#   Date:   22.09.2024             #
####################################
#
#           GitHub Link
# https://github.com/Hin7zer
#
####################################

# Glob al arguments inclusive default values and required flags
Param(
    [Parameter(Mandatory=$false)]
    [switch]$nogui,
    
    [Parameter(Mandatory=$false)]
    [string]$sourcePath,
    
    [Parameter(Mandatory=$false)]
    [string]$exportFile,

    [Parameter(Mandatory=$false)]
    [string]$ignore_FileNames, # will be separated by delimeter ';' later

    [Parameter(Mandatory=$false)]
    [string]$ignore_FileExtensions="csv", # will be separated by delimeter ';' later
    
    [Parameter(Mandatory=$false)]
    [int]$levels = 0

)

# Default global variables
$programm_name = "FileHierarchy2csv"
$programm_version = "Version: 1.0" 
$programm_status = "not started"

# Function for Logging and warnings
function log_message{
    param (
        [string]$log_message, # Log Message
        [string]$log_level, # Loglevel
        [string]$log_file = "", # Logfile Variable
        [string]$log_gui = $false, # Log Message in ui box
        [string]$log_status = $True # Variable if log into file for message
    )
    $timestamp_log = Get-Date -UFormat "%Y-%m-%d %H:%M:%S.%3N" # get timestamp
    $log_level = $log_level.ToUpper() # convert logfile in only upper characters
    $log_line = "[$timestamp_log]:[LOG][NOGUI=$nogui]:[$log_level] $log_message" # the whole log line content
    if ($log_gui -eq $True) {
        [System.Windows.MessageBox]::Show("$log_message","$log_level",0)
    }
    else {
        if ($log_level -eq "WARNING") { # difference output based on log level
            Write-Warning $log_line
        } else {
            Write-Host $log_line
        }
    }

    # Log in file if status is true
    if ($log_status -eq $True) {
        # if log_file is not defined or empty get path of script exec
        if ([string]::IsNullOrEmpty($log_file)) {
            $log_file = Join-Path -Path (Get-Location) -ChildPath "$programm_name.log.txt"
        }

        # Log in file
        Add-Content -Path $log_file -Value $log_line
    }
}

# Function for export tasks
function export_structure {
    ## Function Parameter and default values
    param (
        [string]$sourcePath,
        [string]$exportFile,
        [string]$ignore_FileNames = "", # will be separated by delimeter ';' later
        [string]$ignore_FileExtensions = "" # will be separated by delimeter ';' later
    )
    ## Log/Output Configuration
    log_message -log_level "" -log_message "### Execution of export function started"
    log_message -log_level "INFO"   -log_message "$programm_name in $programm_version started"
    log_message -log_level "DEBUG"  -log_message "Configuration: sourcePath: $sourcePath"
    log_message -log_level "DEBUG"  -log_message "Configuration: exportFile: $exportFile"
    log_message -log_level "DEBUG"  -log_message "Configuration: levels: $levels"
    log_message -log_level "DEBUG"  -log_message "Configuration: ignore_FileNames: $ignore_FileNames"
    log_message -log_level "DEBUG"  -log_message "Configuration: ignore_FileExtensions: $ignore_FileExtensions"

    ## split file filters into array 
    $ignore_FileNames_array = $ignore_FileNames -split ";" | Where-Object { $_.Trim() -ne "" }
    $ignore_FileExtensions_array = $ignore_FileExtensions -split ";" | Where-Object { $_.Trim() -ne "" }

    ## If levels variable is set to 0 (unlimited) the depth will be identified and set automatically
    if ($levels -eq 0) {
        $levels = Get-ChildItem -Path $sourcePath -Recurse -Directory | ForEach-Object {
            ($_).FullName.Substring($sourcePath.Length).TrimStart("\") -split "\\" | Measure-Object | Select-Object -ExpandProperty Count
        } | Sort-Object -Descending | Select-Object -First 1
    }

    ## Get all Files in Source Path and export them into CSV format
    Get-ChildItem -Path $sourcePath -Recurse -File | ForEach-Object {
        ### Set Variables
        $FileName = $_.BaseName
        $Extension = $_.Extension.TrimStart('.')
        $FileFullName = "$FileName.$Extension"

        ### Ignore files based on filters
        if ($ignore_FileNames_array -contains $FileName -or $ignore_FileExtensions_array -contains $Extension) {
            log_message -log_level "INFO" -log_message "ignored File: $sourcePath/$FileFullName"
            return
        }
        
        ### Set variables for file information
        $FullPath = $_.FullName.Substring($sourcePath.Length).TrimStart("\") # Get full path
        $FullPath_splitted = $FullPath -split "\\" # Get full path but splitted
        $FolderPath = $FullPath_splitted[0..($FullPath_splitted.Count - 2)]  # Get directory name
        
        ### Generate object based on files
        $object_export = New-Object PSObject
        
        ### Add filename and extension to object
        Add-Member -InputObject $object_export -MemberType NoteProperty -Name "FileName" -Value $FileName
        Add-Member -InputObject $object_export -MemberType NoteProperty -Name "FileExtension" -Value $Extension

        ### Set levels and add information in object
        for ($i = 1; $i -le $levels; $i++) {
            $levelName = "Level$i"
            $FolderName = if ($FolderPath.Count -ge $i) { $FolderPath[$i - 1] } else { "-" }
            Add-Member -InputObject $object_export -MemberType NoteProperty -Name $levelName -Value $FolderName
        }

        ### return object
        $object_export
    } | Export-Csv -Path $exportFile -NoTypeInformation -Encoding UTF8 # Export information as CSV
}

# Function for integer checks
function IntegerCheck ([string]$intvarcheck){ 
    Try{
        $Null = [convert]::ToInt32($intvarcheck) # Check if Var is integer
        return $true                             # Return true if integer
    }
    Catch{
        log_message -log_level "warning" -log_message "Level Value is not an Integer (Number)!"  # log warning
        return $false                                           # If not integer return false
    }
}

# Function for GUI/UI inclusive Button Functions
function load_gui {

    ## Load libaries for forms function (powershell gui)
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName PresentationCore,PresentationFramework

    ## Window Variables
    $FileBrowser_InitialDirectory = [Environment]::GetFolderPath('Desktop') # Default directory in FileBrowser/FileDialog

    ## Create Variable Object with Gui objects/settings
    $gui = New-Object System.Windows.Forms.Form
    $gui.Text ="$programm_name"
    #$geticon = New-Object system.drawing.icon (".\$programm_name.ico")
    #$gui.Icon = $geticon
    $gui.Width = 500
    $gui.Height = 500

    ## Label settings/placments/attachment (Positions and parameter)
    ### Label Version information
    $label_version = New-Object System.Windows.Forms.Label
    $label_version.Location = New-Object System.Drawing.Size(300,25)
    $label_version.font = New-Object System.Drawing.Font(7,7)
    $label_version.Text = $programm_version
    $label_version.AutoSize = $True
    $gui.Controls.Add($label_version) ## Add label to gui window

    ### Label Author information 
    $label_author = New-Object System.Windows.Forms.Label
    $label_author.Location = New-Object System.Drawing.Size(300,10)
    $label_author.font = New-Object System.Drawing.Font(7,7)
    $label_author.Text = "$programm_name by Benjamin Hinz"
    $label_author.AutoSize = $True
    $gui.Controls.Add($label_author) ## Add label to gui window
    
    ### Label Level information 
    $label_levels = New-Object System.Windows.Forms.Label
    $label_levels.Location = New-Object System.Drawing.Size(10,12)
    $label_levels.font = New-Object System.Drawing.Font(9,9)
    $label_levels.Text = "Levels (Number):"
    $label_levels.AutoSize = $True
    $gui.Controls.Add($label_levels) ## Add label to gui window

    ### Label Level additional information 
    $label_levelsInfo = New-Object System.Windows.Forms.Label
    $label_levelsInfo.Location = New-Object System.Drawing.Size(10,50)
    $label_levelsInfo.font = New-Object System.Drawing.Font(9,9)
    $label_levelsInfo.Text = "How Deep (Levels) should the source folder structure analysed? 0 = automatically"
    $label_levelsInfo.AutoSize = $True
    $gui.Controls.Add($label_levelsInfo) ## Add label to gui window
    
    ### Label Status information 
    $label_status = New-Object System.Windows.Forms.Label
    $label_status.Location = New-Object System.Drawing.Size(200,410)
    $label_status.font = New-Object System.Drawing.Font(12,12)
    $label_status.Text = "Status: $programm_status"
    $label_status.AutoSize = $True
    $gui.Controls.Add($label_status) ## Add label to gui window

    ### Label file name limit information 
    $label_filenameLimit = New-Object System.Windows.Forms.Label
    $label_filenameLimit.Location = New-Object System.Drawing.Size(10,250)
    $label_filenameLimit.font = New-Object System.Drawing.Font(9,9)
    $label_filenameLimit.Text = "Filter for Filenames to ignore them in Export (Delimeter ';'):"
    $label_filenameLimit.AutoSize = $True
    $gui.Controls.Add($label_filenameLimit) ## Add label to gui window

    ### Label file extension limit information 
    $label_fileextLimit = New-Object System.Windows.Forms.Label
    $label_fileextLimit.Location = New-Object System.Drawing.Size(10,300)
    $label_fileextLimit.font = New-Object System.Drawing.Font(9,9)
    $label_fileextLimit.Text = "Filter for file extensions to ignore in export (Delimeter ';'):"
    $label_fileextLimit.AutoSize = $True
    $gui.Controls.Add($label_fileextLimit) ## Add label to gui window

    ### Label Source Path Info
    $label_srcPath = New-Object System.Windows.Forms.Label
    $label_srcPath.Location = New-Object System.Drawing.Size(10,95)
    $label_srcPath.font = New-Object System.Drawing.Font(10,10)
    $label_srcPath.Text = "Source:"
    $label_srcPath.AutoSize = $True
    $gui.Controls.Add($label_srcPath) ## Add label to gui window

    ### Label Export Path Info
    $label_exportPath = New-Object System.Windows.Forms.Label
    $label_exportPath.Location = New-Object System.Drawing.Size(10,175)
    $label_exportPath.font = New-Object System.Drawing.Font(10,10)
    $label_exportPath.Text = "Export Path:"
    $label_exportPath.AutoSize = $True
    $gui.Controls.Add($label_exportPath) ## Add label to gui window

    ### LinkLabel for References - Github/Socielmedia/SourceCode
    $link_github = New-Object System.Windows.Forms.LinkLabel 
    $link_github.Location = New-Object System.Drawing.Size(434,25) 
    $link_github.Size = New-Object System.Drawing.Size(150,20) 
    $link_github.font = New-Object System.Drawing.Font(7,7)
    $link_github.LinkColor = "blue" 
    $link_github.ActiveLinkColor = "blue" 
    $link_github.Text = "GitHub" 
    $link_github.add_Click({[system.Diagnostics.Process]::start("https://github.com/Hin7zer")}) 
    $gui.Controls.Add($link_github) 

    ## Textbox settings/placments/attachment (Positions and parameter)
    ### Textbox Level configuration
    $TextBox_Levels = New-Object System.Windows.Forms.TextBox
    $TextBox_Levels.Location = New-Object System.Drawing.Size(10,30)
    $TextBox_Levels.Size = New-Object System.Drawing.Size(50,500)
    $TextBox_Levels.Text = "$Levels"
    $gui.Controls.Add($TextBox_Levels) ## Add textbox to gui window

    ### Textbox filename limits
    $TextBox_filenameLimit = New-Object System.Windows.Forms.TextBox
    $TextBox_filenameLimit.Location = New-Object System.Drawing.Size(10,275)
    $TextBox_filenameLimit.Size = New-Object System.Drawing.Size(250,10)
    $TextBox_filenameLimit.Text = "$ignore_FileNames"
    $gui.Controls.Add($TextBox_filenameLimit) ## Add textbox to gui window

    ### Textbox file extension limit
    $TextBox_fileextensionLimit = New-Object System.Windows.Forms.TextBox
    $TextBox_fileextensionLimit.Location = New-Object System.Drawing.Size(10,325)
    $TextBox_fileextensionLimit.Size = New-Object System.Drawing.Size(250,10)
    $TextBox_fileextensionLimit.Text = "$ignore_FileExtensions"
    $gui.Controls.Add($TextBox_fileextensionLimit) ## Add textbox to gui window

    ##Button settings/placments/attachment/actions (Positions and parameter)
    ### Button Source Select
    $button_filebrowser_input = New-Object System.Windows.Forms.Button
    $button_filebrowser_input.Location = New-Object System.Drawing.Size(10,120)
    $button_filebrowser_input.Size = New-Object System.Drawing.Size(150,40)
    $button_filebrowser_input.ForeColor  = "white"
    $button_filebrowser_input.BackColor  = "blue"
    $button_filebrowser_input.Text = "Select Source"
    $button_filebrowser_input.Add_Click({ # Button Source Select - Actions/Functions/Commands  - Directory select in dialog
        $FileBrowser_input = New-Object System.Windows.Forms.FolderBrowserDialog
        $FileBrowser_input.Description = "Select a folder"
        $FileBrowser_input.SelectedPath = $FileBrowser_InitialDirectory
    
        if($FileBrowser_input.ShowDialog() -eq "OK")
        {
            $Source = $FileBrowser_input.SelectedPath
            $label_srcPath.Text = "Source: $Source"
            Set-Variable -Name sourcePath -Value $Source -Scope Global
        }
    })

    ### Button Export Select
    $button_filebrowser_output = New-Object System.Windows.Forms.Button
    $button_filebrowser_output.Location = New-Object System.Drawing.Size(10,200)
    $button_filebrowser_output.Size = New-Object System.Drawing.Size(150,40)
    $button_filebrowser_output.ForeColor  = "white"
    $button_filebrowser_output.BackColor  = "blue"
    $button_filebrowser_output.Text = "Select Export-Path"
    $button_filebrowser_output.Add_Click({ # Button Export Select - Actions/Functions/Commands - File select in browser/dialog
        $FileBrowser_output = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
            CheckFileExists = 0  # file does not need to exist (file can be created)
            #ValidateNames = 0  # activate File Validation if neccessary
        }
        $FileBrowser_output.FilterIndex = 1 # Set Filter default
        $FileBrowser_output.filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*" # Filters for file browser/dialog
        $FileBrowser_output.ShowDialog() | Out-Null # open filebrowser/dialog
        $export_file = $FileBrowser_output.FileName # variable 
        $label_exportPath.Text = "Export Path: $export_file" # Set Label text
        Set-Variable -Name exportFile -Value $export_file -Scope Global # Set variable global environment

    })


    $button_actions_export = New-Object System.Windows.Forms.Button
    $button_actions_export.Location = New-Object System.Drawing.Size(10,400)
    $button_actions_export.Size = New-Object System.Drawing.Size(150,40)
    $button_actions_export.ForeColor  = "black"
    $button_actions_export.BackColor  = "green"
    $button_actions_export.Text = "export"
    $button_actions_export.Add_Click({ # Button Export Execution - Actions/Functions/Commands - File select in browser/dialog
        $ignore_FileNames=$TextBox_filenameLimit.Text        # Get file filter into variable
        $ignore_FileExtensions=$TextBox_fileextensionLimit.Text     # Get file filter into variable
        $levels = $TextBox_Levels.Text                       # Get level depth into variable
        $checks_passed = $true                              # default value for checks passed
        $integercheckresponse = IntegerCheck ($levels)      # variable verification in dedicated function | variable with status

        ### Variable verification in condition inclusive Message in dedicated box if failed
        if($integercheckresponse -eq $False){  # If Levels is no integer
            log_message -log_level "info" -log_message "Level Value is not an Integer (Number)! Operation canceled." -log_gui $true -log_status $false # log info in message box
            $checks_passed = $false             # Variable definition for check failed
        }
        if ($sourcePath -eq "") { # If Source is not defined
             log_message -log_level "info" -log_message "Missing Variable, Check Source Path" -log_gui $true -log_status $false # log info in message box
            $checks_passed = $false             # Variable definition for check failed
        }
        if ($exportFile -eq "") { # If Export is not defined
            log_message -log_level "info" -log_message "Missing Variable, Check Export Path" -log_gui $true -log_status $false # log info in message box
            $checks_passed = $false             # Variable definition for check failed
        }
        if ($programm_status -ne "done" -and $programm_status -ne "not started") { # If programm/script status is not as expected i.e. export is running
            log_message -log_level "warning" -log_message "Attention! Process has already started." -log_gui $true -log_status $false # log warning in message box
            $checks_passed = $false
        }
        if ( $checks_passed -eq $false) { # If checks are not passed print info or else condition for start programm/script
            log_message -log_level "info" -log_message "Checks not passed, process will not start" -log_gui $true -log_status $false # log info in message box
            $checks_passed = $false
        }
        else { # Else condition to start programm/script on passed checks
            $programm_status = "running" # set status to running to prevent multiple executions
            $timestamp = Get-Date -UFormat "%Y-%m-%d %R" # get start timestamp in variable
            $label_status.Text = "Status: $programm_status... [$timestamp]" # Set status in label
            export_structure -sourcePath $sourcePath -exportFile $exportFile -levels $levels -ignore_FileNames $ignore_FileNames -ignore_FileExtensions $ignore_FileExtensions # Run export function with parameters
            $programm_status = "done" # set status to done to allow new executions
            $timestamp = Get-Date -UFormat "%Y-%m-%d %R"  # get finished timestamp
            $label_status.Text = "Status: $programm_status! [$timestamp]" # Set status in label
            log_message -log_level "info" -log_message "export is done"  -log_gui $true -log_status $false # log finished export in message box
                
            }
        })
    
 
    ## Add buttons to GUI
    $gui.Controls.Add($button_actions_export) ## Add button to gui window
    $gui.Controls.Add($button_filebrowser_input) ## Add button to gui window
    $gui.Controls.Add($button_filebrowser_output) ## Add button to gui window

    [void]$gui.ShowDialog() # Open script window
}


# Function main is the start function to control which function needs to be loaded
function main {
    ## Check if nogui parameter is set to identify if ui needs to be loaded
    if ($nogui) {
        if (-not $sourcePath -or -not $exportFile) { # Check if neccessary Variables are provided
            log_message -log_level "warning" -log_message "Please provide at least 'sourcePath' and 'exportFile' parameter if you want to use this script without gui." 
            exit # exit script
        }
        export_structure -sourcePath $sourcePath -exportFile $exportFile -levels $levels -ignore_FileNames $ignore_FileNames -ignore_FileExtensions $ignore_FileExtensions # execute function for export inclusive neccessary tasks
        log_message -log_level "info" -log_message "## Script (NOGUI) done"
        exit # exit script
    }        
    else {
    
        ## Load function for GUI/UI
        load_gui 
        log_message -log_level "info" -log_message "### Script (GUI) done/closed"
        exit
    }
}
# Start function main
main
