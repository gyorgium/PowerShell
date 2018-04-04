$apps=@( 	
	#"9E2F88E3.Twitter"
	#"ClearChannelRadioDigital.iHeartRadio"
	#"Flipboard.Flipboard"
	#"king.com.CandyCrushSodaSaga"
	#"Microsoft.3DBuilder"
	#"Microsoft.BingFinance"
	#"Microsoft.BingNews"
	#"Microsoft.BingSports"
	#"Microsoft.BingWeather"
	#"Microsoft.CommsPhone"
	#"Microsoft.Getstarted"
	#"Microsoft.Messaging"
	#"Microsoft.MicrosoftOfficeHub"
	#"Microsoft.MicrosoftSolitaireCollection"
	#"Microsoft.Office.OneNote"
	#"Microsoft.Office.Sway"
	#"Microsoft.People"
	#"Microsoft.SkypeApp"
	#"Microsoft.Windows.Phone"
	#"Microsoft.Windows.Photos"
	#"Microsoft.WindowsAlarms"
	#"Microsoft.WindowsCalculator"
	#"Microsoft.WindowsCamera"
	#"Microsoft.WindowsMaps"
	#"Microsoft.WindowsPhone"
	#"Microsoft.WindowsSoundRecorder"
	#"Microsoft.ZuneMusic"
	#"Microsoft.ZuneVideo"
	#"Microsoft.MinecraftUWP"
	#"ShazamEntertainmentLtd.Shazam"
    #"Microsoft.Windows.FeatureOnDemand.InsiderHub"
    #"Microsoft.Windowscommunicationsapps"
    #"Microsoft.XboxApp"
    #"Microsoft.XboxGameCallableUI"	
    #"Microsoft.XboxIdentityProvider"
    #"Facebook.Facebook"
    #"4DF9E0F8.Netflix"	
)

Function Remove-BuiltInWin10Apps {
    ForEach ($app in $apps) {	
        Write-Host "Removing $app application..."
	    Get-AppxPackage -Name $app -AllUsers | Remove-AppxPackage
	    Get-AppXProvisionedPackage -Online | where DisplayName -eq $app | Remove-AppxProvisionedPackage -Online
        Get-AppXProvisionedPackage -Online | where-object {$_.packagename –like $app} | Remove-AppxProvisionedPackage -Online 
			
	    $appPath="$Env:LOCALAPPDATA\Packages\$app*"
	    Remove-Item $appPath -Recurse -Force -ErrorAction 0
        Write-Host "Done.`n..........................................................................................`n"
    }
    
    Write-Host "Press any key to continue..."
    cmd /c pause | out-null
} # end Remove-BuiltInWin10Apps

Function Add-BuiltInWin10Apps {
    Get-AppxPackage -AllUsers | 
    ForEach-Object { 
        Add-AppxPackage -Register “$($_.InstallLocation)\AppXManifest.xml” -DisableDevelopmentMode 
    } 
} # end Add-BuiltInWin10Apps
