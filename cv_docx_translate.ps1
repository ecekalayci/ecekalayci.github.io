$ErrorActionPreference = 'Stop'

$source = 'C:\Users\ekalayci\Documents\New project\ekalayci_CV_04.2026.docx'
$output = 'C:\Users\ekalayci\Documents\New project\ekalayci_CV_04.2026_ENG_FINAL.docx'
$work = 'C:\Users\ekalayci\Documents\New project\_docx_eng_final'
$zipPath = 'C:\Users\ekalayci\Documents\New project\_docx_eng_final.zip'

if (Test-Path $work) {
  Remove-Item -LiteralPath $work -Recurse -Force
}
if (Test-Path $zipPath) {
  Remove-Item -LiteralPath $zipPath -Force
}
if (Test-Path $output) {
  Remove-Item -LiteralPath $output -Force
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($source, $work)

$xmlPath = Join-Path $work 'word\document.xml'
[xml]$xml = Get-Content -LiteralPath $xmlPath -Raw -Encoding UTF8
$ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')

function Set-ParagraphText {
  param(
    [string]$Old,
    [string]$New
  )

  $matched = $false
  $paras = $xml.SelectNodes('//w:p', $ns)
  foreach ($p in $paras) {
    if ($null -eq $p) { continue }
    $texts = @($p.SelectNodes('.//w:t', $ns))
    if ($texts.Count -eq 0) { continue }
    $current = ($texts | ForEach-Object { $_.'#text' }) -join ''
    if ($current -eq $Old) {
      $texts[0].InnerText = $New
      for ($i = 1; $i -lt $texts.Count; $i++) {
        $texts[$i].InnerText = ''
      }
      $matched = $true
    }
  }
  if (-not $matched) {
    Write-Warning "No paragraph match for: $Old"
  }
}

$replacements = @(
  @('E-posta: ekalayci@pau.edu.tr', 'Email: ekalayci@pau.edu.tr'),
  @('Researchgate Profile: Ece Kalayci', 'ResearchGate Profile: Ece Kalayci'),
  @('Web of Science Profile: ECE KALAYCI - Web of ScienceResearcher Profile', 'Web of Science Profile: ECE KALAYCI - Web of Science Researcher Profile'),
  @('Araştırma Alanları', 'Research Interests'),
  @('Tekstil Malzemeleri ve Prosesleri:', 'Textile Materials and Processes:'),
  @('Yüksek performanslı lifler ve Sürdürülebilir alternatif doğal lifler', 'High-performance fibers and sustainable alternative natural fibers'),
  @('Proses Tasarımı ve Analizi:', 'Process Design and Analysis:'),
  @('Yüksek performanslı liflerin yeni nesil çevreci yöntemlerle boyanması, Sürdürülebilir terbiye teknolojileri, doğal boya/yardımcı kimyasal alternatifleri', 'Dyeing of high-performance fibers using next-generation environmentally friendly methods, sustainable finishing technologies, and natural dye/auxiliary chemical alternatives'),
  @('Odak Alanları:', 'Focus Areas:'),
  @('Sürdürülebilir ve ileri tekstil malzemeleri, ekolojik taşıyıcılar, boyama proseslerinin optimizasyonu ve çevresel etki performansı', 'Sustainable and advanced textile materials, ecological carriers, optimization of dyeing processes, and environmental impact performance'),
  @('Mevcut Kurum', 'Current Position'),
  @('Doktor Öğretim ÜyesiPamukkale Üniversitesi Mühendislik Fakültesi ', 'Assistant Professor, Pamukkale University Faculty of Engineering'),
  @('Tekstil Mühendisliği Bölümü,Tekstil Bilimleri ABD, Denizli, Türkiye', 'Department of Textile Engineering, Division of Textile Sciences, Denizli, Türkiye'),
  @('Şubat, 2026 - Devam', 'February 2026 – Present'),
  @('Araştırma Görevlisi, Dr,Pamukkale Üniversitesi Mühendislik Fakültesi ', 'Research Assistant, Ph.D., Pamukkale University Faculty of Engineering'),
  @('Şubat,2016 -Şubat, 2026 ', 'February 2016 – February 2026'),
  @('Eğitim Bilgileri', 'Education'),
  @('Doktora, Tekstil Mühendisliği, Pamukkale Üniversitesi, Denizli, Türkiye', 'Ph.D., Textile Engineering, Pamukkale University, Denizli, Türkiye'),
  @('Doktora tezi: “Polieterimid liflerinin boyanma özelliklerinin araştırılması”', 'Dissertation: “Polieterimid liflerinin boyanma özelliklerinin araştırılması”'),
  @('Ağustos, 2024', 'August 2024'),
  @('Yüksek Lisans,Tekstil Mühendisliği, Pamukkale Üniversitesi, Denizli, Türkiye', 'M.Sc., Textile Engineering, Pamukkale University, Denizli, Türkiye'),
  @('                 Yüksek Lisans Tezi: ”Ananas liflerinin önterbiyesinin araştırılması”', 'Master''s Thesis: ”Ananas liflerinin önterbiyesinin araştırılması”'),
  @('Ocak, 2017', 'January 2017'),
  @('Lisans, Tekstil Mühendisliği, Pamukkale Üniversitesi, Denizli, Türkiye', 'B.Sc., Textile Engineering, Pamukkale University, Denizli, Türkiye'),
  @('Haziran, 2010', 'June 2010'),
  @('Uluslararası Eğitim ve Akademik Deneyim', 'International Education and Academic Experience'),
  @('Erasmus Öğrenim Hareketliliği', 'Erasmus Study Mobility'),
  @('       Ev Sahibi Kurum:UniversitatPolitecnica de Valencia, Valencia, İspanya', 'Host Institution: Universitat Politècnica de València, Valencia, Spain'),
  @('Eylül 2008-Haziran 2009', 'September 2008 – June 2009'),
  @('Uluslararası Akademik Ziyaret / Araştırma Deneyimi', 'International Academic Visit / Research Experience'),
  @('       Ev Sahibi Kurum: Kyoto Institute of Technology, Kyoto, Japonya', 'Host Institution: Kyoto Institute of Technology, Kyoto, Japan'),
  @('Temmuz,2016- Ağustos,2016', 'July 2016 – August 2016'),
  @('Desteklenen Projeler', 'Funded Projects'),
  @('2026 – Devam ediyor', '2026 – Ongoing'),
  @('Colorsense -Boyali Kumaş Renk Kalite Kontrolü Süreci İyileştirilmesi İçin Yapay Zeka Destekli Karar Sistemi Tasarimi, Projedeki Görevi: Araştırmacı/uzman, TUBITAK 1505, 2025-2027.', 'Colorsense -Boyali Kumaş Renk Kalite Kontrolü Süreci İyileştirilmesi İçin Yapay Zeka Destekli Karar Sistemi Tasarimi, Role in Project: Researcher/Expert, TUBITAK 1505, 2025-2027.'),
  @('2025 – Devam ediyor', '2025 – Ongoing'),
  @('Denizli’de yaygın olarak üretilen bitkilerin polyester liflerinin renklendirilmesinde doğal difüzyon arttırıcı olarak kullanımı, Proje yürütücüsü, Başlangıç Seviyesi Projesi, Pamukkale Üniversitesi, 2025BSP006.', 'Denizli’de yaygın olarak üretilen bitkilerin polyester liflerinin renklendirilmesinde doğal difüzyon arttırıcı olarak kullanımı, Principal Investigator, Starter-Level Project, Pamukkale University, 2025BSP006.'),
  @('Polieterimid liflerinin boyanma özelliklerinin araştırılması, Projedeki görevi: Araştırmacı, Doktora Tez Projesi, Tamamlandı, 2024.', 'Polieterimid liflerinin boyanma özelliklerinin araştırılması, Role in Project: Researcher, Ph.D. Thesis Project, Completed, 2024.'),
  @('Ananas liflerinin ön terbiyesinin araştırılması, Projedeki görevi: Araştırmacı,', 'Ananas liflerinin ön terbiyesinin araştırılması, Role in Project: Researcher,'),
  @('Yüksek Lisans Tez Projesi, Tamamlandı, 2016', 'Master''s Thesis Project, Completed, 2016'),
  @('Yayınlar', 'Publications'),
  @('SCI / SCI-Expanded makaleler', 'SCI / SCI-Expanded Articles'),
  @('Uluslararası Hakemli Dergilerde Yayınlanan Makaleler (Alan İndeksli Dergiler)', 'Articles Published in International Peer-Reviewed Journals (Field-Indexed Journals)'),
  @('Ulusal hakemli dergilerde yayınlanan makaleler', 'Articles Published in National Peer-Reviewed Journals'),
  @('Uluslararası kitap ve kitap bölümleri', 'International Books and Book Chapters'),
  @('Uluslararası bildiriler', 'International Conference Papers'),
  @(' (Ayrıca makale olarak basılmıştır: ', ' (Also published as: '),
  @(' (Ayrıcamakaleolarakbasılmıştır: ', ' (Also published as: '),
  @('Ulusal bildiriler', 'National Conference Papers'),
  @('Hakemlikler', 'Reviewer Activities'),
  @('Tekstil veKonfeksiyon', 'Tekstil ve Konfeksiyon'),
  @('Tekstil veMühendis', 'Tekstil ve Mühendis'),
  @('Harran ÜniversitesiMühendislikDergisi', 'Harran Üniversitesi Mühendislik Dergisi'),
  @('Yabancı Diller', 'Languages'),
  @('İngilizce (İleri), İspanyolca (Başlangıç)', 'English (Advanced), Spanish (Beginner)'),
  @('Dersler', 'Courses Taught'),
  @('2025-2026 Bahar Dönemi, TENG 455, Mühendislik Etiği ', 'Spring 2025-2026, TENG 455, Engineering Ethics'),
  @('2025-2026 Bahar Dönemi, TENG 613, DigitalTransformation in TextileandArtificialIntelligence Applications,  ', 'Spring 2025-2026, TENG 613, Digital Transformation in Textile and Artificial Intelligence Applications')
)

foreach ($pair in $replacements) {
  Set-ParagraphText -Old $pair[0] -New $pair[1]
}

$xml.Save($xmlPath)

Compress-Archive -Path (Join-Path $work '*') -DestinationPath $zipPath -Force
Move-Item -LiteralPath $zipPath -Destination $output

Write-Output $output
