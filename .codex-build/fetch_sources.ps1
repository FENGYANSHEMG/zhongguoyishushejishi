$ErrorActionPreference = 'Stop'
$assetRoot = Join-Path $PSScriptRoot 'research-assets'
New-Item -ItemType Directory -Path $assetRoot -Force | Out-Null
$entries = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'sources.json') -Raw | ConvertFrom-Json
$entries | ForEach-Object -Parallel {
  $entry = $_
  $outRoot = $using:assetRoot
  try {
    $htmlPath = Join-Path $outRoot ($entry.id + '.html')
    $page = Invoke-WebRequest -Uri $entry.url -TimeoutSec 25
    [System.IO.File]::WriteAllText($htmlPath, $page.Content)
    $pics = @($page.Images | Where-Object { $_.src -match '(?i)\.(jpg|png|jpeg)(\?|$)' } | ForEach-Object {
      [PSCustomObject]@{url=([uri]::new([uri]$entry.url,$_.src)).AbsoluteUri; alt=$_.alt}
    } | Sort-Object -Property url -Unique)
    $ordered = @($page.Images | Where-Object { $_.src -match '(?i)\.(jpg|png|jpeg)(\?|$)' } | ForEach-Object {
      [PSCustomObject]@{url=([uri]::new([uri]$entry.url,$_.src)).AbsoluteUri; alt=$_.alt}
    })
    $chosen = $null
    if ($entry.url -match 'chnmuseum.cn/zp') { $chosen = $ordered | Where-Object url -Match '/[PW]0[0-9]+\.(jpg|png)' | Select-Object -First 1 }
    elseif ($entry.url -match 'www.dpm.org.cn/collection') { $chosen = $ordered | Where-Object { $_.alt -and $_.alt -ne '故宫博物院' -and $_.url -match '/Uploads/(Picture|image)/' } | Select-Object -First 1 }
    elseif ($entry.id -eq 'R25') { $chosen = $ordered | Where-Object url -Match 'shuziwenwu' | Select-Object -First 1 }
    if ($chosen) {
      $ext = [IO.Path]::GetExtension(([uri]$chosen.url).AbsolutePath)
      $imgPath = Join-Path $outRoot ($entry.id + $ext)
      Invoke-WebRequest -Uri $chosen.url -OutFile $imgPath -TimeoutSec 25
      $entry | Add-Member imagePath $imgPath -Force
      $entry | Add-Member imageUrl $chosen.url -Force
    }
    $entry | Add-Member images $pics -Force
    $plain = $page.Content -replace '(?is)<script\b[^>]*>.*?</script>','' -replace '(?is)<style\b[^>]*>.*?</style>','' -replace '<[^>]+>',' ' -replace '\s+',' '
    $entry | Add-Member text ([System.Net.WebUtility]::HtmlDecode($plain)) -Force
    [IO.File]::WriteAllText((Join-Path $outRoot ($entry.id + '.json')), ($entry | ConvertTo-Json -Depth 5))
    "$($entry.id) OK image=$([bool]$chosen)"
  } catch { "$($entry.id) ERROR $($_.Exception.Message)" }
} -ThrottleLimit 5
