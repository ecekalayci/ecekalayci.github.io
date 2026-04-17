$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.IO.Compression.FileSystem

$source = 'C:\Users\ekalayci\Documents\New project\ekalayci_CV_04.2026.docx'
$output = 'C:\Users\ekalayci\Documents\New project\ekalayci_CV_04.2026_ENG_FINAL.docx'
$work = 'C:\Users\ekalayci\Documents\New project\_docx_eng_final2'
$zipPath = 'C:\Users\ekalayci\Documents\New project\_docx_eng_final2.zip'

if (Test-Path $work) { Remove-Item -LiteralPath $work -Recurse -Force }
if (Test-Path $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
if (Test-Path $output) { Remove-Item -LiteralPath $output -Force }

[System.IO.Compression.ZipFile]::ExtractToDirectory($source, $work)
$xmlPath = Join-Path $work 'word\document.xml'
[xml]$xml = Get-Content -LiteralPath $xmlPath -Raw -Encoding UTF8
$ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
$paras = @($xml.SelectNodes('//w:p',$ns))

function SetPara([int]$n, [string]$new) {
  $p = $paras[$n-1]
  if ($null -eq $p) { throw "Missing paragraph $n" }
  $texts = @($p.SelectNodes('.//w:t',$ns))
  if ($texts.Count -eq 0) { throw "Paragraph $n has no text nodes" }
  $texts[0].InnerText = $new
  for ($i = 1; $i -lt $texts.Count; $i++) {
    $texts[$i].InnerText = ''
  }
}

$map = [ordered]@{
  4 = "Department of Textile Engineering,"
  5 = "Faculty of Engineering, Pamukkale University"
  9 = "Email: ekalayci@pau.edu.tr"
  13 = "ResearchGate Profile: Ece Kalayci"
  15 = "Web of Science Profile: ECE KALAYCI - Web of Science Researcher Profile"
  18 = "Research Interests"
  20 = "Textile Materials and Processes:"
  21 = "High-performance fibers and sustainable alternative natural fibers"
  22 = "Process Design and Analysis:"
  23 = "Dyeing of high-performance fibers using next-generation environmentally friendly methods, sustainable finishing technologies, and natural dye/auxiliary chemical alternatives"
  24 = "Focus Areas:"
  25 = "Sustainable and advanced textile materials, ecological carriers, optimization of dyeing processes, and environmental impact performance"
  27 = "Current Position"
  29 = "Assistant Professor, Pamukkale University Faculty of Engineering"
  30 = "Department of Textile Engineering, Division of Textile Sciences, Denizli, Türkiye"
  31 = "February 2026 – Present"
  35 = "Research Assistant, Ph.D., Pamukkale University Faculty of Engineering"
  36 = "Department of Textile Engineering, Division of Textile Sciences, Denizli, Türkiye"
  37 = "February 2016 – February 2026"
  40 = "Education"
  42 = "Ph.D., Textile Engineering, Pamukkale University, Denizli, Türkiye"
  43 = "Dissertation: Investigation of the Dyeing Properties of Polyetherimide Fibers"
  45 = "August 2024"
  46 = "M.Sc., Textile Engineering, Pamukkale University, Denizli, Türkiye"
  47 = "Master's Thesis: Investigation of the Pretreatment of Pineapple Fibers"
  49 = "January 2017"
  50 = "B.Sc., Textile Engineering, Pamukkale University, Denizli, Türkiye"
  52 = "June 2010"
  54 = "International Education and Academic Experience"
  56 = "Erasmus Study Mobility"
  57 = "Host Institution: Universitat Politècnica de València, Valencia, Spain"
  59 = "September 2008 – June 2009"
  60 = "International Academic Visit / Research Experience"
  61 = "Host Institution: Kyoto Institute of Technology, Kyoto, Japan"
  62 = "July 2016 – August 2016"
  64 = "Funded Projects"
  67 = "2026 – Ongoing"
  72 = "Colorsense -Boyali Kumaş Renk Kalite Kontrolü Süreci İyileştirilmesi İçin Yapay Zeka Destekli Karar Sistemi Tasarimi, Role in Project: Researcher/Expert, TUBITAK 1505, 2025-2027."
  73 = "2025 – Ongoing"
  75 = "Denizli’de yaygın olarak üretilen bitkilerin polyester liflerinin renklendirilmesinde doğal difüzyon arttırıcı olarak kullanımı, Principal Investigator, Starter-Level Project, Pamukkale University, 2025BSP006."
  76 = "2025 – Ongoing"
  79 = "Polieterimid liflerinin boyanma özelliklerinin araştırılması, Role in Project: Researcher, Ph.D. Thesis Project, Completed, 2024."
  82 = "Ananas liflerinin ön terbiyesinin araştırılması, Role in Project: Researcher,"
  83 = "Master's Thesis Project, Completed, 2016"
  86 = "Publications"
  88 = "SCI / SCI-Expanded Articles"
  96 = "Articles Published in International Peer-Reviewed Journals (Field-Indexed Journals)"
  102 = "Articles Published in National Peer-Reviewed Journals"
  113 = "International Books and Book Chapters"
  122 = "International Conference Papers"
  151 = "Avinç, O. O., Yıldırım, F. F., Yavaş, A. & Kalayci, E., (2017, May). 3D printingtechnologyanditsinfluences on thetextileindustry. In IIER International Conference. Beijing, China. (Also published as: Avinc, O., Yildirim, F. F., Yavas, A., & Kalayci, E. (2017). 3D printingtechnologyanditsinfluences on thetextileindustry. International Journal of IndustrialElectronicsandElectricalEngineering, 5(7), 37. http://iraj.in)"
  152 = "Kalayci, E., Yavaş, A., & Avinç, O. (2018, November 22–23). The effects of different alkali treatments with different temperatures on the colorimetric properties of lignocellulosic raffia fibers. In International Conference on Recent Advances in Engineering and Technology. Havana, Cuba. (Also published as: Kalayci, E., Avinc, O., & Yavas, A. (2019). The effects of different alkali treatments with different temperatures on the colorimetric properties of lignocellulosic raffia fibers. International Journal of Advances in Science Engineering and Technology, 7(Special Issue 1), 15. http://iraj.in"
  158 = "National Conference Papers"
  165 = "Reviewer Activities"
  169 = "Tekstil ve Konfeksiyon"
  172 = "Tekstil ve Mühendis"
  174 = "Harran Üniversitesi Mühendislik Dergisi"
  176 = "Languages"
  178 = "English (Advanced), Spanish (Beginner)"
  180 = "Courses Taught"
  182 = "Spring 2025-2026, TENG 455, Engineering Ethics"
  183 = "Spring 2025-2026, TENG 613, Digital Transformation in Textile and Artificial Intelligence Applications"
}

foreach ($k in $map.Keys) {
  SetPara -n $k -new $map[$k]
}

$xml.Save($xmlPath)
Compress-Archive -Path (Join-Path $work '*') -DestinationPath $zipPath -Force
Move-Item -LiteralPath $zipPath -Destination $output
Write-Output $output
