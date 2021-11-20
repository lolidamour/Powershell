#Bar de progression dans une boucle Foreach
#Avant la boucle Foreach il faut placer et initialiser une variable $i = 0
Function ProgressBar {
  param([String]$ListCount,
      [String]$TitreProgress)
      
  $PercentComplete = [System.Math]::Round($($i*100/($ListCount)),2)
  Write-Progress -Activity $TitreProgress -status "Effectué : $PercentComplete %" -percentcomplete $($i*100/($ListCount))
  $i++
}
