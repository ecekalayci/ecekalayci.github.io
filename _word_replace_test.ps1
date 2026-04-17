$ErrorActionPreference = 'Stop'
$src = 'C:\Users\ekalayci\Documents\New project\ekalayci_CV_04.2026.docx'
$out = 'C:\Users\ekalayci\Documents\New project\_word_test.docx'
$wdFindContinue = 1
$wdReplaceAll = 2
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
  $doc = $word.Documents.Open($src)
  $doc.SaveAs2($out)
  $find = $doc.Content.Find
  $find.ClearFormatting()
  $find.Replacement.ClearFormatting()
  $find.Text = 'Researchgate Profile'
  $find.Replacement.Text = 'ResearchGate Profile'
  $find.Forward = $true
  $find.Wrap = $wdFindContinue
  $find.Format = $false
  $find.MatchCase = $false
  $find.MatchWholeWord = $false
  $find.MatchWildcards = $false
  $find.MatchSoundsLike = $false
  $find.MatchAllWordForms = $false
  [void]$find.Execute($find.Text,$find.MatchCase,$find.MatchWholeWord,$find.MatchWildcards,$find.MatchSoundsLike,$find.MatchAllWordForms,$find.Forward,$find.Wrap,$find.Format,$find.Replacement.Text,$wdReplaceAll)
  $doc.Save()
  $doc.Close()
}
finally {
  $word.Quit()
}
Write-Output $out
