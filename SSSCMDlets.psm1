#.ExternalHelp SSSCMDlets.psm1-help.xml
Function Export-ToExcel{
<#
.NAME
   ExportTo-Excel
.AUTHOR
    Jonathan Bailey
.SYNOPSIS
   Allows user to choose a CSV file to export to his local machine.
.SYNTAX
   ExportTo-Excel [-InitialDirectory] <String[]> [-ExportFolder] <String[]> [-Description] <String[]> [-DisplayName] <String[]>
.DESCRIPTION
   When the function is called, an open file dialog box appears prompting the user to choose a file.
   Once a file is chosen, a second box appears prompting the user to select a folder to ouput to.
   This CMDlet is designed to be run from a script that is custom configured for each user.  The -DisplayName
   parameter MUST BE UNIQUE to that person's session or else the CMDlet will not work.
.EXAMPLE
   ExportTo-Excel -InitialDirectory D:\RPG -ExportFolder \\tsclient\c\output -Description "Steve's Transfer" -DisplayName "Steve"
   
   This command directs the user to the "D:\RPG" folder so that he may choose a file for export.  The folder browser
   windows will open up with \\tsclient\c\output as its root and the user will be allowed to choose what folder he
   wants to put the file in.  The transfer job will be called "Steve" and its description is "Steve's Transfer".
#>
    [CmdletBinding(DefaultParameterSetName="Transfer",
                   SupportsShouldProcess=$true,
                   PositionalBinding=$true
                  )]
    Param(
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $InitialDirectory,
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        $ExportFolder,
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        $Description,
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=3)]
        $DisplayName
    )
    BEGIN{
        $date = get-date -UFormat %m-%d-%y

        [void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
        [void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.window")

        # Create an open file dialog box.  User finds csv file to use for data import.
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.title = "Please select the file created for export:"
        $OpenFileDialog.initialDirectory = $InitialDirectory
        $OpenFileDialog.filter = "Text File (*.TXT)| *.TXT"
        $OpenFileResult = $OpenFileDialog.ShowDialog()
        $TXT = $OpenFileDialog.filename
        if ($OpenFileResult -eq [System.Windows.Forms.DialogResult]::Cancel){
            throw "Program Cancelled."
        }

        if ((Test-Path -Path $ExportFolder) -eq $false){
            New-Item -ItemType Directory -Path $ExportFolder
        }

        # Creates a progress bar
        $Title = "BITS Transfer Progress"
        # winform dimensions
        $height=100
        $width=400
        # winform background color
        $color = "White"

        # create the form
        $form1 = New-Object System.Windows.Forms.Form
        $form1.Text = $title
        $form1.Height = $height
        $form1.Width = $width
        $form1.BackColor = $color

        $form1.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle 
        # display center screen
        $form1.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

        # create label
        $label1 = New-Object system.Windows.Forms.Label
        $label1.Text = "not started"
        $label1.Left=5
        $label1.Top= 10
        $label1.Width= $width - 20
        # adjusted height to accommodate progress bar
        $label1.Height=15
        $label1.Font= "Verdana"
        # optional to show border 
        #$label1.BorderStyle=1

        # add the label to the form
        $form1.controls.add($label1)

        $progressBar1 = New-Object System.Windows.Forms.ProgressBar
        $progressBar1.Name = 'progressBar1'
        $progressBar1.Value = 0
        $progressBar1.Style="Continuous"

        $System_Drawing_Size = New-Object System.Drawing.Size
        $System_Drawing_Size.Width = $width - 40
        $System_Drawing_Size.Height = 20
        $progressBar1.Size = $System_Drawing_Size

        $progressBar1.Left = 5
        $progressBar1.Top = 40
        $form1.Controls.Add($progressBar1)
    }

    PROCESS{
        $exportfolder = Join-Path -Path $ExportFolder -ChildPath $date
        if ((Test-Path -Path $ExportFolder) -eq $false){
            New-Item -ItemType Directory -Path $ExportFolder
        }
        Try{
            Import-Module BitsTransfer
            Start-BitsTransfer -Destination $ExportFolder -Source $TXT -Description $Description -DisplayName $DisplayName -Priority Normal -Asynchronous -RetryInterval 60 -RetryTimeout 120 -ErrorAction Stop -ErrorVariable $exporterror

            $form1.Show()| out-null

            #give the form focus
            $form1.Focus() | out-null

            #update the form
            $label1.Text = "Preparing to send files"
            $form1.Refresh()

            start-sleep -Seconds 1

            $bits = Get-BitsTransfer -Name $DisplayName
            $pct = 0
            while ($bits.JobState -ne "Transferred"  -and $pct -ne 100){
                if ($bits.jobstate -eq "Error" -or $bits.JobState -eq "TransientError" ){
                    Resume-BitsTransfer -BitsJob $bits
                }
   
                $pct = ($bits.BytesTransferred / $bits.BytesTotal)*100
                $progressbar1.Value = $pct
                Start-Sleep -Milliseconds 100
                $label1.text="Sending file..."
                $form1.Refresh()
                    }

                    $form1.Close()
            if($bits.jobstate -eq "Transferred"){
                [System.Windows.Forms.MessageBox]::Show("$DisplayName was transferred successfully.  Please check your Output folder to open the file.")
            }
        }
        Catch{
        
        Get-LogError -LogErrorVariable "ExportError" -LogFolder c:\log -LogFile "ExportTo-ExcelError.txt"
        }
    }
    END{
        Remove-BitsTransfer -BitsJob $DisplayName -ErrorAction Stop -ErrorVariable exporterror
    }
}
#.ExternalHelp SSSCMDlets.psm1-help.xml
function Get-LogError{
<#
.NAME
    Get-LogError
.AUTHOR
    Jonathan Bailey
.SYNOPSIS
   Gets errors from error variable $logerrorvariable and puts them into a text file that buffers up to 300 lines. 
.DESCRIPTION
   When this function is called, it looks for what's in $logerror and outputs everything in the array.
.SYNTAX
    Get-LogError [-LogErrorVariable] <String[]> [-LogFolder] <String[]> [-LogFile] <String[]>
.EXAMPLE
   get-logerror -LogErrorVariable "logerror" -LogFolder c:\log -LogFile "ExportTo-ExcelLog.txt"
   
   "logerror" can be anything you want the errorvariable to be.
   However, this must be used in conjunction with the common parameter -errorvariable in
   whatever command you're trying to log errors from.
#>
    [CmdletBinding(DefaultParameterSetName="CreateLog",
                   SupportsShouldProcess=$true,
                   PositionalBinding=$true)]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $LogErrorVariable,
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        $LogFolder,
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        $LogFile
    )

    Begin{
       
        if((Test-Path $LogFolder) -eq $false){
            md $LogFolder
        }
        
        $log = Join-Path -Path $LogFolder -ChildPath $LogFile
        $buffer = get-content $log
        if((test-path $log) -eq $true -and $buffer.Length -gt 300){
            Clear-Content $log
            $buffer | select first 300 | Set-Content $log
        }
        if((Test-Path $log) -eq $false){ 
            New-Item -Path $LogFolder -Name $LogFile -ItemType file -Force
        }
        $i = 0
    }
    Process{
        ForEach ($l in $LogErrorVariable){
            ((get-date).DateTime + "-" + $LogErrorVariable[$i]) | Add-Content $log
            $i++
        }
    }
    End{}
}
#.ExternalHelp SSSCMDlets.psm1-help.xml
function Install-Chocolatey{
    [CmdletBinding(DefaultParameterSetName ="Install",
                   SupportsShouldProcess=$true,
                   PositionalBinding=$false)]
    Param()
    BEGIN{
        $dir = $ENV:HOMEDRIVE
        $date = (get-date).DateTime
        $error.Clear()
    }
    PROCESS{
        if (test-path ("$dir\nugetlog.txt")){
            Clear-Content -Path $dir\nugetlog.txt
        }
        Add-Content -Path $dir\nugetlog.txt -Value "$date - Nuget Script has attempted to run on this machine."

        # tests to see if chocolatey is installed on the machine.
        if((Test-Path ("$env:ProgramData\chocolatey")) -eq $false){
            Add-Content -path $dir\nugetlog.txt -Value "$date - Chocolatey not found.  Attempting install."
            try{
                # execution already set to remotesigned.  This should run.
                # Set-ExecutionPolicy unrestricted -Scope Process -force
                iex ((new-object net.webclient).DownloadString('https://chocolatey.org/install.ps1')) -ErrorAction Stop
                Add-Content -Path $dir\nugetlog.txt -Value "$date - Chocolatey installed successfully. $error"
            }
            catch{
                if (!($error[0] -eq $null)){
                    Add-Content -Path $dir\nugetlog.txt -Value "$date - $error[0].fullyqualifiederrorid"
                }
            }
        }
        else{
            Add-Content -Path $dir\nugetlog.txt -Value "$date - Chocolatey is already installed on this machine.  Updating..."
            cup # chocolatey update.  Updates to the latest package of chocolatey.

            # use this only for testing.
            # write-host "Test is good."
        }
    }
    END{}
}
#.ExternalHelp SSSCMDlets.psm1-help.xml
Function Install-WMF3{
    [CmdletBinding(DefaultParameterSetName="Install",
                   SupportsShouldProcess=$true,
                   PositionalBinding=$false)]
    Param()

    BEGIN{}
    PROCESS{
        Try{
    
            if ((Get-HotFix -Id kb2506146 -ErrorAction stop -ErrorVariable $logerror).hotfixid -notcontains "KB2506146" `
                -or (Get-HotFix -id kb2506143 -ErrorAction stop -ErrorVariable $logerror).hotfixid -notcontains "KB2506143" `
                -and (gwmi win32_operatingsystem -ErrorAction stop -ErrorVariable $logerror).Version -contains "6.0"){
                    $logerror += choco install wmf3 -version 3.0.20121027 -y
 
            }
            else{
                throw "WMF 3.0 Can't be installed.  Machine must be Kernel 6.0, or WMF 3.0 is already installed."
                $logerror += $error[0]           
            }
        }
        Catch{
            Get-LogError $logerror
        }
    }
    END{
        $logerror.clear()
    }

}
#.ExternalHelp SSSCMDlets.psm1-help.xml
Function Install-WMF4{
    [CmdletBinding(DefaultParameterSetName="Install",
                   SupportsShouldProcess=$true,
                   PositionalBinding=$false)]
    Param()

    BEGIN{}
    PROCESS{
        Try{
            if ((Get-HotFix -id KB2819745 -ErrorAction stop -ErrorVariable $Logerror).hotfixid -notcontains "KB2819745" `
                -and (gwmi win32_operatingsystem -ErrorAction stop -ErrorVariable $logerror).version -contains "6.1"){
                    $logerror += choco install powershell -y
            }
            else{
                throw "WMF 4.0 Can't be installed.  Machine must be Kernel 6.1, or WMF 4.0 is already installed."
                $Logerror += $error[0]
            }
        }
        Catch{
            Get-LogError $Logerror
        }
    }
    END{
        $Logerror.clear()
    }

}
#.ExternalHelp SSSCMDlets.psm1-help.xml
Function Install-ReportViewer11{
    [CmdletBinding(DefaultParameterSetName="Install",
                   SupportsShouldProcess=$true,
                   PositionalBinding=$false)]
    Param()

    BEGIN{}
    PROCESS{
        choco install reportviewer.2012 -version 11.1.3452.0 -y
    }
    END{}

}
#.ExternalHelp SSSCMDlets.psm1-help.xml
Function Install-SQLCLRTypes{
    [CmdletBinding(DefaultParameterSetName="Install",
                   SupportsShouldProcess=$true,
                   PositionalBinding=$false)]
    Param()

    BEGIN{}
    PROCESS{
        Try{
            if((gwmi -Class win32_product -ErrorAction stop -ErrorVariable $logerror).name -ne "Microsoft System CLR Types for SQL Server 2012"){
                $logerror += choco install sql2012.clrtypes -y
            }
            else{
                throw "SQL CLR Types is already installed on this system."
                $logerror += $error[0]
            }
        }
        Catch{
            Get-LogError $logerror
        }
    }
    END{
        $logerror.clear()
    }

}
#.ExternalHelp SSSCMDlets.psm1-help.xml
Function Install-DotNet45{
    [CmdletBinding(DefaultParameterSetName="Install",
                   SupportsShouldProcess=$true,
                   PositionalBinding=$false)]
    Param()
    BEGIN{}
    PROCESS{
        choco install dotnet4.5 -y
    }
    END{}
}
#.ExternalHelp SSSCMDlets.psm1-help.xml
Function Install-MapPoint{
    [CmdletBinding(DefaultParameterSetName="Install",
                   SupportsShouldProcess=$true,
                   PositionalBinding=$true)]
    Param( 
           [Parameter(Mandatory=$true,
                     ValueFromPipelineByPropertyName=$true,
                     Position=0)]
           $dir       
          )
    BEGIN{
        $test = Get-WmiObject -Class win32_product | Where-Object {
            $_.name -eq "Microsoft Mappoint North America 2006"
        }
    }
    PROCESS{
        if ( $test.name -ne "Microsoft Mappoint North America 2006"){
            msiexec /i $dir\map2006\mappoint\msmap\data.msi /qn /norestart /le c:\mappointlog.txt OFFICE_INTEGRATION=0
        }
    }
    END{}
}
#.ExternalHelp SSSCMDlets.psm1-help.xml
Function Set-UEFIActivation{
<#
.NAME
    Set-UEFIActivation
.AUTHOR
	Jonathan Bailey
.SYNOPSIS
    Pulls the Product key hex value from the motherboard, converts it to ASCII, and pipes it to slmgr.vbs for activation.
.SYNTAX
   Set-UEFIActivation [-OA3ToolPath] <string> 
.DESCRIPTION
    This script pulls the product key from a UEFI/Windows 8 motherboard for use in oem key activation in legacy mode.
    Make sure to Have oa3tool.exe in the same working directory as this script.  It is part of Microsoft ADK.
    Use this script to install 32-bit versions of Windows 8.1 on UEFI OEM machines running in legacy mode.  
    Since activation doesn't occur automatically, manual activation is necessary.
#>
    [CmdletBinding(DefaultParameterSetName="OA3ToolPath",
                   SupportsShouldProcess=$true,
                   PositionalBinding=$true
                  )]
    Param(
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $OA3ToolPath
    )
]
    BEGIN{
        $Arg = "/validate"
    }
    PROCESS{
        # pull the hex value from motherboard and outputs it to $hexdata
        $HexData = "$OA3ToolPath $Arg"

        # Find the hex value that contains the product key and formats/trims it for conversion.
        $HexData = $HexData | select -First 33 | select -Last 4
        $HexData = $HexData -replace '\s+', ' '
        $HexData = $HexData.trimstart(' ')
        $HexData = $HexData.trimend(' ')

        # Split hex values into objects and convert them to decimal, then decimal to ASCII, 
        # then set the new value as $prodkey.
        $HexData.split(" ") | FOREACH {[CHAR][BYTE]([CONVERT]::toint16($_,16))} | Set-Variable -name prodkey -PassThru

        # join the ascii array into a string
        $prodkey = $prodkey -join ''
        # regex replace all unprintable characters.
        $prodkey = $prodkey -replace "[^ -x7e]",""

        write-host
        write-host success!

        # use slmgr.vbs for activation.
        slmgr /ipk $prodkey
    }
    END{}
}
#.ExternalHelp SSSCMDlets.psm1-help.xml
Function New-AVException{
<#
.NAME
    New-AVException
.AUTHOR
	Jonathan Bailey
.SYNOPSIS
   Generates a list of UNC paths to add to exclusion / exceptions lists for Antivirus.
.SYNTAX
    New-AVException
.DESCRIPTION
    This script will attempt to get psdrive informationand if an error occurs, it will
    attempt to gather the information using net use.  If everything goes well, a list of exceptions
    will be piped into a gridview to copy into the AV exceptions field.
#>

# Gives proper formatting to error status for Get-PsDrive
    [CMDletBinding(DefaultParameterSetName="DisplayExceptions",
                   SupportsShouldProcess=$true,
                   PositionalBinding=$false)]
    Param()
    BEGIN{
        $error.clear()
        $drivemap = $null
        $errornull = @"
Error status is Null. Using Get-PsDrive to generate UNC Path.
"@
    }
    PROCESS{
        # Attempts to generate UNC Data using Get-PsDrive
        try {
            $test = Get-PSDrive -Name I -ErrorAction Stop
            if ($test.DisplayRoot -notlike "\\*"){
                throw "DisplayRoot does not contain a UNC path."
            }
    
            if (!$error){
                write-host $errornull -ForegroundColor Yellow
                $drivemap = Get-PSDrive -Name I | Select-Object DisplayRoot | ft -HideTableHeaders
                # Converts powershell display object to a string and trims whitespace.
                $drivemap = $drivemap | Out-String
                $drivemap = $drivemap.trim()
            }
        }

        # If Get-PsDrive was not used to create the drive map, Net Use is utilized instead.
        catch {
            if ($error[0].FullyQualifiedErrorId -eq "GetLocationNoMatchingDrive,Microsoft.Powershell.Commands.GetPSDriveCommand" -or $error[0].FullyQualifiedErrorId -eq "DisplayRoot does not contain a UNC Path."){
                Write-Host "$error  Using Net Use to generate UNC path." -ForegroundColor Yellow
                $drivemap = net use I:
                #selects the line to use and removes all unnecessary expressions.
                $drivemap = $drivemap | select -Skip 1 -First 1
                $drivemap = $drivemap.trim("Remote name ")
            }

        }

        # Creates an array with the child paths and uses foreach to attach them to the UNC that's been generated.
        finally{
            $array = @(
                "\rpg\",
                "\rpg\vblib",
                "\rpg\#library",
                "\rpg\#library\b36run.exe",
                "\rpg\vblib\propane.exe",
                "\rpg\vblib\update.exe",
                "\rpg\#library\oclrt.exe",
                "\rpg\#library\wsio.exe",
                "\rpg\#library\wsiostop.exe"
                )

            $exceptions = @()

            $array | foreach {

                $exceptions += Join-Path -Path $drivemap -ChildPath $_ 
        
            }

            # Sends the array to the Grid Viewer to copy into the exceptions list.
            $exceptions | Out-GridView -Title "UNC Path Exceptions"
        }
    }
    END{}
}
#.ExternalHelp SSSCMDlets.psm1-help.xml
Function Set-DriveMaps{
<#
.NAME
    Set-DriveMaps
.AUTHOR
	Jonathan Bailey
.SYNOPSIS
    sets drivemaps to a file share not on the domain through secure string credentials.
.SYNTAX
   Set-driveMaps [-Username] <String[]> [-Password] <String[]> 
.DESCRIPTION
    when run as part of a script, this allows users to log in using custom credentials.
    It maps all drive required for internal SSS mission critical file access.
.EXAMPLE
   Set-DriveMaps -Username "Contoso\administrator" -Password "P@ssw0rd"
   
   This CMDlet allows an administrator to quickly set drive maps using individual usernames and passwords to a
   file share that does not have domain trust.  It works best when using GPO Mapped drives is not optimal, like
   say, when users need to use individual logins different than their domain credentials to access a network 
   drive.
#>
    [CMDletbinding(DefaultParameterSetName="Credentials",
                   SupportsShouldProcess=$true,
                   PositionalBinding=$true
                   )
    ]
    Param(
          [Parameter(Mandatory=$true,
                     ValueFromPipelineByPropertyName=$true,
                     Position=0)]
          $Username,
          [Parameter(Mandatory=$true,
                     ValueFromPipelineByPropertyName=$true,
                     Position=1)]
          $Password    
    )
    BEGIN{
        $pass= $Password|ConvertTo-SecureString -AsPlainText -Force
        $Cred = New-Object System.Management.Automation.PsCredential($Username,$pass)
        $logname = "drivemaplog.txt"
        $dir = "c:\log"
        $log = join-path -Path $dir -ChildPath $logname
    
        if((test-path $dir) -eq $false){
            md c:\log
        }
        if((test-path $log) -eq $false){
            new-item -Name $logname -Path $dir -ItemType file
        }

        $drivearray = @()
        $driveletter = @("I","O","T","S","U")
        $name = @("APPS","ClIENTELE","DOCS","STORAGE","USERS")
        $date = get-date -UFormat %m-%d-%y
        $i = 0
    }
    PROCESS{
        foreach ($d in $driveletter){
            $drive = New-Object -TypeName system.object
            $drive | Add-Member -MemberType NoteProperty -Name "DriveLetter" -Value $d
            $drive | Add-Member -MemberType NoteProperty -Name "Name" -Value $name[$i]
            $drive | Add-Member -MemberType NoteProperty -Name "UNCPath" -Value (join-path -path \\IBM2008 -childpath $name[$i])
            $drivearray += $drive
            $i++
        }

        $drivearray | foreach {

            try{
                # use this for testing
                # net use $_.driveletter $_.uncpath 
                net use ($_.driveletter+":")
                if ($LASTEXITCODE -eq "0"){
                   $del = net use ($_.driveletter+":") /delete
                   Add-Content -Value "$date - $del" -Path $log
                }
                if ((get-psdrive -Name ($_.driveletter) -ErrorAction SilentlyContinue) -eq ($_.driveletter)){
                    Remove-PSDrive -Name ($_.driveletter) -ErrorVariable pserror -ErrorAction Stop
                }
            }
            catch{
                if ($LASTEXITCODE -ne "0"){
                    Add-Content -Value "$date - $error[0].fullyqualifiederrorid - $LASTEXITCODE" -Path $log
                    if($pserror[0].fullyqualifiederrorid -eq "DriveNotFound,Microsoft.PowerShell.Commands.RemovePSDriveCommand"){
                        Add-Content -Value "$date - $pserror[0].fullyqualifiederrorid" -Path $log
                    }
                }
            }
            Finally{
                try{
                    New-PSDrive -Name ($_.driveletter) -PSProvider FileSystem -Root ($_.uncpath) -Description "($_.name)" -Scope Global -Persist -Credential $cred -ErrorVariable finallyerror
                }
                catch{
                    if ($finallyerror){
                        Add-Content -Value "$date - $finallyerror[0].fullyqualifiederrorid" -Path $log
                    }
                }
            }
        }
    }
    END{}
}
#.ExternalHelp SSSCMDlets.psm1-help.xml
Function Rename-SSSPC{
    <#
    .AUTHOR
        Jonathan Baily and Tyler Brown
    .SYNOPSIS 
       Renames a computer. 
    .DESCRIPTION
       Renames the computer by grabing the serial number off the motherboard and replacing the current computer name.
    .EXAMPLE
       Rename-SSSPC
    .INPUTS
       No inputs required.  This CMDlet pulls information from the motherboard of the PC.
    .OUTPUTS
       No outputs.  The data is automatically piped to the Rename-Computer CMDlet.
    .NOTES
       This CMDlet requires no parameters.  This CMDlet only runs on PowerShell 4.0.
    .COMPONENT
       This CMDlet belongs to the SSSCMDlets Module.
    #>
    [CMDletbinding( DefaultParameterSetName = 'Static', 
                    SupportsShouldProcess=$true, 
                    PositionalBinding=$false)]
    Param()
    BEGIN {}
    PROCESS {
        $serial = gcim win32_bios | select -Property serialnumber | ft -HideTableHeaders
        $serial = $serial | out-string
        $serial = $serial.trim( )
        Rename-Computer -ComputerName (gcim win32_operatingsystem).CSName -NewName $serial
    }
    END{}
}
#.ExternalHelp SSSCMDlets.psm1-help.xml
Function New-ShortcutIcons{
<#
.NAME
    New-Shortcuts
.AUTHOR
    Jonathan Bailey
.SYNOPSIS
    Creates shortcuts corresponding to the proper workstation IDs.
.DESCRIPTION
    Can either pull shortcut IDs from a CSV, or from a set of IDs provided as a parameter.
.EXAMPLE
    New-Shortcuts -IDName AA,A1,A2 -RPGDirectory "I:\RPG" -Version "SSS" -Destination "I:\icons"
.EXAMPLE
    New-Shortcuts -CSV "D:\icons.csv" -RPGDirectory "D:\RPG" -Version "Propane" -Destination "D:\icons"
.INPUTS
    Can either take indivdual or multiple WSIDs, or a CSV file with WSIDs.  If using a CSV, please label
    the column WSID.
.OUTPUTS
    Creates a shortcut link in the provided destination folder.
.COMPONENT
    This CMDlet is part of SSSCMDlets Module.
#>
    [CmdletBinding(DefaultParameterSetName ='WSID', 
                   SupportsShouldProcess=$true, 
                   PositionalBinding=$false
                  )]
    Param
    (
        # Used in declaring WSID manually.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='WSID')]
        $IDName,

        [Parameter( Mandatory = $true,
                    ValueFromPipeline = $true,
                    ValueFromPipelineByPropertyName = $true,
                    Position = 1
                  )]
        $RPGDirectory,

        # Choose SSS or Propane
        [Parameter( Mandatory = $true,
                    ValueFromPipeline = $false,
                    ValueFromRemainingArguments = $false,
                    Position = 2
                  )]
        [ValidateSet("SSS","Propane")]
        $Version,

        [Parameter( Mandatory = $true,
                    ValueFromPipeline = $true,
                    ValueFromRemainingArguments = $false,
                    Position = 3
                  )]
        [String]
        $Destination,

        # Param3 help description
        [Parameter( ParameterSetName='CSV',
                    Mandatory = $true,
                    ValueFromPipeline = $true,
                    ValueFromRemainingArguments = $false,
                    Position = 4
                  )]
        [String]
        $CSV
    )

    Begin{
        $WSID = @()
        Set-Location $Destination
    }
    Process{
        If ($CSV.Length -gt 0){
            $CSVar = Import-Csv $CSV
            $CSVar | foreach {
                $WSID += $_.WSID
            }
        }
        If ($IDName.Length -gt 0){
            $IDName | foreach {
                $WSID += $_
            }
        }
        If($Version -eq "SSS"){
            $WorkDir = Join-Path $RPGDirectory -ChildPath "#library"
            $RPGProgram = Join-Path $WorkDir -ChildPath "b36run.exe"
            $IconName = "SSS"
        }
        If ($Version -eq "Propane"){
            $WorkDir = Join-Path $RPGDirectory -ChildPath "VBLib"
            $RPGProgram = Join-Path $WorkDir -ChildPath "propane.exe"
            $IconName = "Propane"
        }

        # Create a Shortcut with Windows PowerShell
        $i = 0
        $WSID | foreach {
            $TargetFile = $RPGProgram
            $ShortcutFile = $Destination + "\" + $IconName + " " + $WSID[$i] + ".lnk"
            $WScriptShell = New-Object -ComObject WScript.Shell
            $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
            $Shortcut.TargetPath = $TargetFile
            $Shortcut.Arguments = $WSID[$i]
            $Shortcut.WorkingDirectory = $WorkDir
            $Shortcut.Save()
            $i++
        }
    }
    END{}
}
#.ExternalHelp SSSCMDlets.psm1-help.xml
Function New-GitRepo{
<#
.NAME
   New-GitRepo
.AUTHOR
   Jonathan Bailey
.SYNOPSIS
   Creates a Git repository at file path the user specifies.
.SYNTAX
   New-GitRepo [-RepoName] <String[]> [-GitRepo] <String[]>
.DESCRIPTION
   Function checks to see if git is installed, and if it isn't it installs git.  If git is installed,
   the function then runs git to make a clone of the primary repository located on github.
.EXAMPLE
   New-GitRepo -RepoName ssscmdlets -GitRepo https://github.com/contoso/sss-cmdlets.git
   
   This command creates a repository at $home\modules\ssscmdlets 
#>

    [CMDletBinding(DefaultParameterSetName="CreateRepo",
                   SupportsShouldProcess=$true,
                   PositionalBinding=$true
                  )]

    Param(

    [Parameter(Mandatory=$true,
               ValueFromPipelineByPropertyName=$true,
               Position=0)]
    $RepoName,
    [Parameter(Mandatory=$true,
               ValueFromPipelineByPropertyName=$true,
               Position=0)]
    $GitRepo
    )
    BEGIN{
        $RepoDir = Join-Path -Path $home\Documents\WindowsPowerShell\Modules\ -ChildPath $RepoName
    }

    PROCESS{

        $GitPath = Join-Path ${env:ProgramFiles(x86)} -ChildPath \git
        if (test-path ($GitPath) -eq $false){
            choco.exe install git
        }


        if (test-path ($RepoDir) -eq $false){
            New-Item -Path $RepoDir -ItemType directory
        }

        Set-Location $RepoName
        git.exe init
        git.exe clone $GitRepo
    }
    END{}
}