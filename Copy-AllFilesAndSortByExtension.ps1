function Copy-AllFilesAndSortByExtension {
    param (
        [Parameter(Mandatory = $true, Position = 0, HelpMessage="Folder to analyze.")]
        [string]$SourcePath,
        [array]$ExtenstionsArray = @("jpg", "jpeg", "png", "tif", "tiff", "pdf", "doc", "docx", "xls", "xlsx", "ppt", "pptx", "csv", "txt", "mp4", "m4a", "mp3")
    )

    # verify that source file exists
    if (!(Test-Path $SourcePath)) {
        Write-Error -Message "Destination $SourcePath was not found!"
    } else {
        # create list of existing files
        $FilesList = New-Object System.Collections.Generic.List[System.Object]
    
        # iterate over all files by extention
        foreach ($Extension in $ExtenstionsArray) {
        
            # get all files by extension
            Get-ChildItem -Path $SourcePath -Recurse -Filter "*.$Extension" | ForEach-Object {$FilesList.Add($_)}

            # if documents found
            if ($FilesList.Count -gt 0) {
                # extraction paths
                $ExtractionPath = "extracted-files\$Extension"
                $ExtractionDuplicatesPath = "$ExtractionPath\dulpicates"
            
                # create extraction path if does not exist
                if (!(Test-Path $ExtractionPath)) {
                    mkdir -Path $ExtractionPath -Force
                }

                # iterate and copy all files found
                foreach ($File in $FilesList) {
                    # copy file if does not exist on destination
                    if (!(Test-Path "$ExtractionPath\$File")) {
                        Copy-Item -Path $File.FullName -Destination "$ExtractionPath\$File" -Force
                    } else {
                        # create duplicates folder if does not exist
                        if (!(Test-Path $ExtractionDuplicatesPath)) {
                            mkdir -Path $ExtractionDuplicatesPath
                        }
                        
                        # copy duplicates to separate folder
                        $DuplicatePrefix = New-Guid
                        Copy-Item -Path $File.FullName -Destination "$ExtractionDuplicatesPath\$DuplicatePrefix+$File" -Force
                    }
                }
            } else {
                Write-Warning -Message "No *.$Extension files found!"
            }
        
            # clear the list before jumping to another extension
            $FilesList.Clear()        
        }
    }    
}

# copy all files
Copy-AllFilesAndSortByExtension

# prompt to finish script
Read-Host -Prompt "Press Enter to continue"