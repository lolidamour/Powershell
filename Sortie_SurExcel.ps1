# Création de l'objet application Excel sinon on réalise un export au format CSV
try{
    $objExcel = new-object -comobject excel.application
    Write-Host "Excel is installed on this Computer, disabled Users will be export in a fashioned excel file."
    $ExcelTest = $true
}catch{
    Write-Host "Excel is not installed on this Computer, disabled Users will be export in a plain old CSV file."
    $ExcelTest = $false
}

# Création d'un nouveau fichier
        $finalWorkBook = $objExcel.Workbooks.Add()
        $finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
        $finalWorkBook.Worksheets.Item(1).Name = "DisabledAccounts"
        #$objExcel.Visible =$true
 
 Write-Host "Create header"
 
 # Rempli la première ligne
    $finalWorkSheet.Cells.Item(1,1) = "TitreColone1"
    $finalWorkSheet.Cells.Item(1,1).Font.Bold = $True ## Si on veut changer le texte en gras
    $finalWorkSheet.Cells.Item(1,2) = "TitreColone2"
    $finalWorkSheet.Cells.Item(1,2).Font.Bold = $True
    $finalWorkSheet.Cells.Item(1,3) = "TitreColone3"
 
 Write-Host "Riceving data..." -ForegroundColor Green
 
 # On commence à la seconde ligne (la 1ère est consacrée au Header)
$FinalExcelRow = 2
# Pour le Choix d'une couleur pour surligner la ligne
$ColorIndex = 41 #Bleu
 
 $Script =...
 
 Write-Host "Writing data..." -ForegroundColor Green

foreach($rusult in $script){
#On stocke les différentes valeurs
        $finalWorkSheet.Cells.Item($FinalExcelRow,1) = $Result1
        #On attribut la couleur définit plus haut pour la case concerné
        $finalWorkSheet.Cells.Item($FinalExcelRow,1).Interior.ColorIndex = $ColorIndex
        $finalWorkSheet.Cells.Item($FinalExcelRow,2) = $Result2
        $finalWorkSheet.Cells.Item($FinalExcelRow,2).Interior.ColorIndex = $ColorIndex
        $finalWorkSheet.Cells.Item($FinalExcelRow,3) = $Result3
        $finalWorkSheet.Cells.Item($FinalExcelRow,3).Interior.ColorIndex = $ColorIndex

# On incrémente le numéro de la ligne en cours d'écriture pour paser à la suivante
       $FinalExcelRow++
}

Write-Host "Saving data and closing Excel." -ForegroundColor Green
$finalWorkBook.SaveAs($ExcelPath)

 #On ferme le fichier
    $finalWorkBook.Close()
    
# Le processus Excel utilisé pour traiter l'opération est arrêté
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)
