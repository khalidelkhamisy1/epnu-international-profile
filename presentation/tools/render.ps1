# Renders every slide of the deck to PNG via PowerPoint COM (true rendering, real fonts).
param(
  [string]$Deck = "D:\AI\EPNU\presentation\EPNU_International_Profile.pptx",
  [string]$OutDir = "D:\AI\EPNU\presentation\render",
  [int]$Width = 1600
)
New-Item -ItemType Directory -Force $OutDir | Out-Null
Get-ChildItem $OutDir -Filter "slide-*.png" -ErrorAction SilentlyContinue | Remove-Item -Force
$pp = New-Object -ComObject PowerPoint.Application
try {
  $p = $pp.Presentations.Open($Deck, $true, $false, $false)  # read-only, untitled, no window
  $h = [int]($Width * 9 / 16)
  foreach ($s in $p.Slides) {
    $f = Join-Path $OutDir ("slide-{0:D2}.png" -f $s.SlideIndex)
    $s.Export($f, "PNG", $Width, $h)
  }
  "Rendered $($p.Slides.Count) slides to $OutDir"
  $p.Close()
} finally { $pp.Quit() }
