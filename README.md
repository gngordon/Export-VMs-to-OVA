Export VMs to OVA

Assists with exporting virtual machines to OVA files

GUI that provides:
	* Selection of VMs from a list from a comma separated file that contains the VM names.
	* Selection of destination directory.

USAGE
.\exportvmtoova.ps1 [vmlist.csv] [vCenterUser] [vCenterPassword]

WHERE
vmlist.csv       = Comma delimited file with a VM per row. Fields required is: Name
vCenterUser      = Username for vCenter Server.
vCenterPassword  = Password for vCenter Server user.

EXAMPLES
.\exportvmtoova.ps1
.\exportvmtoova.ps1 mylist.csv
.\exportvmtoova.ps1 mylist.csv administrator@vsphere.local VMware1!

REQUIREMENTS
Requires PowerCLI.
Change variable for vCenter.
List of VMs in comma separated file.