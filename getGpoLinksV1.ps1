[CmdletBinding()] 
    param (
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)][string[]]$ous,
		[Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)][string[]]$gpoNames,
		$path = "",
		$filter = { 1 -eq 1 }
    )
#initialize the output variable
$results = @{ "Linked" = @(); "Inherited" = @(); "OU" = @(); }; 	

#if ous were passed in, get the linked policies
if($ous -ne $null -and $ous -ne ""){
	#get the referenced OU
	foreach($ou in $ous){
		write-verbose "Parsing OU: '$ou'"
		Get-ADOrganizationalUnit -filter { DistinguishedName -eq $ou} -Properties name,distinguishedName,gpLink,gPOptions |  % {
			#get the gpLinks property from that OU
			$gpLinks = $_.gPlink.split("][")            
			$gpLinks =  @($gpLinks | ? {$_})            
			if($gpLinks -ne $null){
				$links = @()
				
				
				#loop through each linked GPO and add it to the array
				foreach($GpLink in $gpLinks) {
					# if($gpLink.count -gt 2){
					
						#get GPO Properties
						$GpName = [adsi]$GPlink.split(";")[0] | select -ExpandProperty displayName            
						write-verbose "Getting details on $GpName"
						$GpStatus = $GPlink.split(";")[1]            
						$EnableStatus = $EnforceStatus = 0
						$gpo = Get-GPO -Name $GPName | select GpoStatus, CreationTime, ModificationTime

						switch($GPStatus) {            
							"1" {$EnableStatus = $false; $EnforceStatus = $false}            
							"2" {$EnableStatus = $true; $EnforceStatus = $true}            
							"3" {$EnableStatus = $false; $EnforceStatus = $true}            
							"0" {$EnableStatus = $true; $EnforceStatus = $false}            
						}            
						
						#if it is valid, add it to the array
						if($GPName -ne $null){
							$links += New-Object PSObject -Property @{
								"Type" = "LinkedGpo";
								"Path" = $ou;
								"DisplayName" = $GPName;
								"Enabled" = [bool]$($EnableStatus);
								"Enforced" = [bool]$($EnforceStatus);
								"Status" = $gpo.gpoStatus;
								"Created" = $gpo.CreationTime;
								"Modified" = $gpo.ModificationTime;
							}
						}
					# }
				}
				$results.Linked += $links
			}
			write-verbose "Analyzing Inherited GPOs"
			
			#now get the inherited GPOs
			$inherit = @()
			#for each inherited GPO for the selected OU
			Get-GPInheritance -Target $ou | select  -expand inheritedGpoLinks | % {
				#if it is valid and not in the CURRENT ou (Get-GPInheritance gets inherited and linked GPOs.  boo)
				if($_.DisplayName -ne $null -and ( $links | select -expand DisplayName ) -notContains $_.DisplayName ){
					#get GPO Properties
					$gpo = Get-GPO -Name $_.DisplayName | select GpoStatus, CreationTime, ModificationTime
					write-verbose "Getting details on $($_.DisplayName)"
					
					#add it to the inherited array
					$inherit += New-Object PSObject -Property @{
						"Type" = "InheritedGpo";
						"Path" = $ou;
						"DisplayName" = $_.DisplayName;
						"Enabled" = [bool]$($_.Enabled);
						"Enforced" = [bool]$($_.Enforced);
						"Status" = $gpo.gpoStatus;
						"Created" = $gpo.CreationTime;
						"Modified" = $gpo.ModificationTime;
					}
				}
			}
			$results.Inherited += $inherit
		}
	}
}

if($gpoNames -ne $null -and $gpoNames -ne ""){
	foreach ($n in $gpoNames) {
		write-verbose "Parsing GPO: '$n'"
		$problem = $false 
		try { 
			Write-Verbose -Message "Attempting to produce XML report for GPO: $n" 
			[xml]$report = Get-GPOReport -Name $n -ReportType Xml -ErrorAction Stop 
		} catch { 
			$problem = $true 
			Write-Warning -Message "An error occurred while attempting to query GPO: $n" 
		} 
		if (-not($problem)) { 
			
			$linksTo = @()
			$report.GPO.LinksTo | % {
				$gpo = Get-GPO -Name $n | select GpoStatus, CreationTime, ModificationTime
				
				$linksTo += New-Object PSObject -Property @{
					"Type" = "LinkedOU"
					"Path" = ((( $_.SOMPath -split "/")[100..1] | % { "OU=$_" }) -join ",") + "," + ((($_.SOMPath -split "/")[0] -split "\." | % {"DC=$($_)"} ) -join "," )
					"Enabled" = [bool]$($_.Enabled);
					"Enforced" = [bool]$($_.NoOverride);
					"DisplayName" = $n;
					"Status" = $gpo.gpoStatus;
					"Created" = $gpo.CreationTime;
					"Modified" = $gpo.ModificationTime;
					
				}
			}
			$results.OU += $linksTo
		}
		
		
	} 
}

$rawOutput = @()

$results.Linked | sort-object -property Path | % { $rawOutput += $_ }
$results.Inherited | sort-object -property Path | % { $rawOutput += $_ }
$results.OU | sort-object -property Path | % { $rawOutput += $_ }

if($path){
	$rawOutput | ? $filter | select Type, Path, DisplayName, Status,  Enabled, Enforced, Created, Modified | export-csv -noTypeInformation $path
}else{
	$rawOutput | ? $filter | select Type, Path, DisplayName, Status,  Enabled, Enforced, Created, Modified
}
