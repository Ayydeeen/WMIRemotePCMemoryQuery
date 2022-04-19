$Workstations = Get-Content "C:\workstations.txt"
$Array = @()

ForEach ($Station in $Workstations) {
    $Check = $Processor = $ComputerMemory = $Object = $null

    Try {
        $PCMemory = (Get-WmiObject -ComputerName $Station  -Class win32_operatingsystem -ErrorAction Stop)
        $Memory = ((($PCMemory.TotalVisibleMemorySize - $PCMemory.FreePhysicalMemory)*100)/ $PCMemory.TotalVisibleMemorySize)
        $RoundMemory = [math]::Round($Memory, 2)

        $Object = New-Object PSCustomObject

        $Object | Add-Member -MemberType NoteProperty -Name "PC name" -Value $Station
        $Object | Add-Member -MemberType NoteProperty -Name "Memory %" -Value $RoundMemory
        
        $Object
        $Array += $Object
    }
    Catch {
        Write-Host "Something went wrong ($station): "$_.Exception.Message
        Continue
    }
}

If ($Array) {
    $Array | Out-GridView -Title "Memory Util"
    $Array | Export-Csv -Path "C:\results.csv" -NoTypeInformation -Force
}
