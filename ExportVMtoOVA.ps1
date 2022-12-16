<#
.SYNOPSIS
Script to Export selected VMs to OVA files.
Requires PowerCLI.
Settings.ini = Contains the variables for vCenter. Change as required.
List of VMs in comma separated file. (defaults to vmlist.csv)

.USAGE
     .\exportvmtoova.ps1 [vmlist.csv] [vCenterUser] [vCenterPassword]
     
     WHERE
         vmlist.csv       = Comma delimited file with a VM per row. Fields required are: Name
         vCenterUser      = Username for vCenter Server.
         vCenterPassword  = Password for vCenter Server user.

.EXAMPLES
     .\exportvmtoova.ps1
     .\exportvmtoova.ps1 mylist.csv
     .\exportvmtoova.ps1 mylist.csv administrator@vsphere.local VMware1!

.ACTIONS  
    *Select VMs
    *Select destination directory
    *Export VMs to OVA
    	
.NOTES
    Version:        2.0
    Author:         Graeme Gordon - ggordon@vmware.com
    Creation Date:  2022/12/16
    Purpose/Change: Export VMs to OVA
  
    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL 
    VMWARE,INC. BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
    IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
    CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 #>

param([string]$vmListFile = "vmlist.csv", [string] $vCenterUser, [string] $vCenterPassword)

#region variables
################################################################################
#                                    Variables                                 #
################################################################################
$SettingsFile			= "settings.ini"

$global:vcConnected		= $false
#endregion variables

function ImportIni
{
################################################################################
# Function to parse token values from a .ini configuration file                #
################################################################################
	param ($file)

	$ini = @{}
	switch -regex -file $file
	{
            "^\s*#" {
                continue
            }
    		"^\[(.+)\]$" {
        		$section = $matches[1]
        		$ini[$section] = @{}
    		}
    		"([A-Za-z0-9#_]+)=(.+)" {
        		$name,$value = $matches[1..2]
        		$ini[$section][$name] = $value.Trim()
    		}
	}
	$ini
}

function Initialize_Env  ($vCenterServer)
{
################################################################################
#               Function Initialize_Env                                        #
################################################################################
    # --- Initialize PowerCLI Modules ---
    #Get-Module -ListAvailable VMware* | Import-Module
    Import-Module VMware.VimAutomation.Core
	Set-PowerCLIConfiguration -Scope User -ParticipateInCeip $false -InvalidCertificateAction ignore -DefaultVIServerMode Multiple -Confirm:$false

    # --- Connect to the vCenter server ---
	$attempt = 0
    Do {
        Write-Output "", "Connecting To vCenter Server:"
		Write-Host ("Connecting To vCenter Server: " + $vCenterServer) -ForegroundColor Yellow
        If (!$vCenterUser)
        {
            $vc = Connect-VIServer -Server $vCenterServer -ErrorAction SilentlyContinue
        }
        elseif (!$vCenterPassword)
        {
            $vc = Connect-VIServer -Server $vCenterServer -User $vCenterUser -ErrorAction SilentlyContinue
        }
        else
        {
             $vc = Connect-VIServer -Server $vCenterServer -User $vCenterUser -Password $vCenterPassword -Force -ErrorAction SilentlyContinue
        }
        If (!$vc.IsConnected)
		{
			$attempt += 1
			Write-Host ("Failed to connect to vCenter Server. Attempt " + $attempt + " of 3")  -ForegroundColor Red
		}
    } Until ($vc.IsConnected -or $attempt -ge 3)
	If ($vc.IsConnected) { $global:vcConnected = $true }
}

function Get_Folder ($Initial_Directory)
{
################################################################################
#             Function Get_Folder                                              #
################################################################################
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select a folder"
    $foldername.rootfolder = "MyComputer"
    $foldername.SelectedPath = $Initial_Directory
    
    if($foldername.ShowDialog() -eq "OK")
    {
        $folder = $foldername.SelectedPath
    }
    return $folder
}

function Define_GUI
{
################################################################################
#              Function Define_GUI                                             #
################################################################################
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $global:form                     = New-Object System.Windows.Forms.Form
    $form.Text                       = 'Export VMs to OVA'
    $form.Size                       = New-Object System.Drawing.Size(500,400)
    #$form.Autosize                   = $true
    $form.StartPosition              = 'CenterScreen'
    $form.Topmost                    = $true

    #OK button
    $OKButton                        = New-Object System.Windows.Forms.Button
    $OKButton.Location               = New-Object System.Drawing.Point(300,320)
    $OKButton.Size                   = New-Object System.Drawing.Size(75,23)
    $OKButton.Text                   = 'OK'
    $OKButton.DialogResult           = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton               = $OKButton
    $form.Controls.Add($OKButton)

    #Cancel button
    $CancelButton                    = New-Object System.Windows.Forms.Button
    $CancelButton.Location           = New-Object System.Drawing.Point(400,320)
    $CancelButton.Size               = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text               = 'Cancel'
    $CancelButton.DialogResult       = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    #Browse for Directory button
    $BrowseButton                    = New-Object System.Windows.Forms.Button
    $BrowseButton.Location           = New-Object System.Drawing.Point(300,40)
    $BrowseButton.Size               = New-Object System.Drawing.Size(150,23)
    $BrowseButton.Text               = 'Select Destination'
    $BrowseButton.Add_Click({ $global:Export_Directory = Get_Folder $Export_Directory })     
    $form.Controls.Add($BrowseButton)

    #Checkbox on whether to demo actions
    $global:DemoSelect                = New-Object System.Windows.Forms.CheckBox
    $DemoSelect.Location              = New-Object System.Drawing.Point(300,80)
    $DemoSelect.Size                  = New-Object System.Drawing.Size(200,23)
    $DemoSelect.Text                  = 'Demo'    
    $DemoSelect.Checked               = $Demo
    $DemoSelect.Add_CheckStateChanged({ $global:Demo = $DemoSelect.Checked })
    $form.Controls.Add($DemoSelect) 

    #Text above list box of VMs
    $label                          = New-Object System.Windows.Forms.Label
    $label.Location                 = New-Object System.Drawing.Point(10,20)
    $label.Size                     = New-Object System.Drawing.Size(250,20)
    $label.Text                     = 'Select VMs from the list below:'
    $form.Controls.Add($label)

    #List box for selection of VMs
    $global:listBox                 = New-Object System.Windows.Forms.Listbox
    $listBox.Location               = New-Object System.Drawing.Point(10,40)
    $listBox.Size                   = New-Object System.Drawing.Size(260,250)
    $listBox.Height                 = 250
    $listBox.SelectionMode          = 'MultiExtended'
    ForEach ($vm in $vmlist)
    {
        [void] $listBox.Items.Add($vm.Name)
    }
    $form.Controls.Add($listBox)
}

#region main
################################################################################
#              Main
################################################################################
Clear-Host

#Check the settings file exists
if (!(Test-path $SettingsFile)) {
	WriteErrorString "Error: Configuration file ($SettingsFile) not found."
	Exit
}
#Import settings variables
$global:vars = ImportIni $SettingsFile
If ($vars.Controls.Demo = "No") { $global:Demo = $False } Else { $global:Demo = $True }
$global:Export_Directory = $vars.Export.Export_Directory

#Check the VM list file exists
if (!(Test-path $vmListFile)) {
	WriteErrorString "Error: VM list file ($vmListFile) not found."
	Exit
}
$global:vmlist = Import-Csv $vmListFile #Import the list of VMs

Define_GUI
$result = $form.ShowDialog()
if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    #Write-Host ("OK Button Pressed") -ForegroundColor Green
    $selection = $listBox.SelectedItems   

    If ($selection)
    {
        Write-Host ("Selected VMs     : " + $selection) -ForegroundColor Yellow
        Write-Host ("Export Directory : " + $Export_Directory) -ForegroundColor Yellow
        Write-Host ("Demo             : " + $Demo) -ForegroundColor Green

        Initialize_Env $vars.vSphere.vCenterServer
		If (!$vcConnected) { Exit }
        Add-Type -AssemblyName 'PresentationFramework'
            
        ForEach ($vm in $vmlist)
        {
            If ($selection.Contains($vm.Name))
            {
                $vmrecord = Get-VM -Name $vm.Name -ErrorAction SilentlyContinue
                if ($vmrecord)
                {
                    Write-Host ("Export VM: " + $vm.Name) -ForegroundColor Green
                    If (!$Demo) { $vmrecord | Export-VApp -Destination $Export_Directory -Format Ova -Force }
                }
                else
                {
                    Write-Host ("VM does not exist: " + $vm.Name) -ForegroundColor Red
                }
            }
        }
    }
    else
    {
        Write-Host ("No VMs Selected") -ForegroundColor Yellow
    }
}
else
{
    #Write-Host ("Cancel Button Pressed") -ForegroundColor Red
}
#endregion logic