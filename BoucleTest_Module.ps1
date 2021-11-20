
# Test de la présence du module Active Directory avec Try/Catch
$ModuleName = ""

if((Get-Module $ModuleName) -eq $null){
    try{
        Import-Module $ModuleName
    }catch{
        Write-Host "The execution computer doesn't have $ModuleName Powershell Module. The script can't continue." -ForegroundColor Red
        return
    }
}

#Exemple : Module Active Directory
if((Get-Module ActiveDirectory) -eq $null){
    try{
        Import-Module ActiveDirectory
    }catch{
        Write-Host "The execution computer doesn't have ActiveDirectory Powershell Module. The script can't continue." -ForegroundColor Red
        return
    }
}
