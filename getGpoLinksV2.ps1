[CmdletBinding()] 
    param (
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)][string[]]$ous,
		[Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)][string[]]$gpoNames,
		$path = "",
		$filter = { 1 -eq 1 }
    )

begin{
	class GpoLinks{
		[String[]]$results = @();
		[String] $execPath;
		[String[]]$ous = (Get-Variable -Name 'ous' -ErrorAction SilentlyContinue).value;
		[String[]]$gpoNames = (Get-Variable -Name 'gpoNames' -ErrorAction SilentlyContinue).value;
		[String]$path = (Get-Variable -Name 'path' -ErrorAction SilentlyContinue).value;
		$filter = (Get-Variable -Name 'filter' -ErrorAction SilentlyContinue).value;
		$self;
		
		GpoLinks(){
			$this.execPath = $PSScriptRoot;
			$this.self = $this
		}
		
		[void] parseOu($ou){
			Get-ADOrganizationalUnit -filter { DistinguishedName -eq $ou} -Properties name,distinguishedName,gpLink,gPOptions |  % {
				$gpLinks = $_.gPlink.split("][")            
				$gpLinks =  @($gpLinks | ? {$_})            
				$links = @()
				
				
				foreach($GpLink in $gpLinks) {
					$GpName = [adsi]$GPlink.split(";")[0] | select -ExpandProperty displayName            
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
				}
				$this.results += $links
				
			}
		}
		
		[void] readOus(){
			if($this.ous -ne $null -and $this.ous -ne ""){
				foreach($ou in $this.ous){
					$this.parseOu($ou);
					
				}
			}
		}

		[String[]] output(){
			$this.results.getType()
			if($this.path){
				$this.results | sort-object -property Path, Type | ? $this.filter | select Type, Path, DisplayName, Status,  Enabled, Enforced, Created, Modified | export-csv -noTypeInformation $this.path
				return $null
			}else{
				$this.results | ft | out-string | write-host
				return $this.results | sort-object -property Path, Type | ? $this.filter | select Type, Path, DisplayName, Status,  Enabled, Enforced, Created, Modified 
			}
			
		}
		
		[void] Dispose(){
			
		}
	}
}

process{
	$global:gpoLinks = [GpoLinks]::new()
	$global:gpoLinks.readOus()	
}

end{
	$global:gpoLinks.output();
	$global:gpoLinks.Dispose();
	[System.GC]::Collect() | out-null
}
