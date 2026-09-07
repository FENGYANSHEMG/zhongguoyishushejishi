$ErrorActionPreference='Stop'
$assetRoot=Join-Path $PSScriptRoot 'research-assets'
$extra=@{
 R27='https://www.nxkg.org.cn/uploadfile/2019/0131/20190131093450138.jpg'
 R32='https://www.publishingmuseum.net/upload/HistoricalStory/202401/51498110965a12a6204671.jpg'
 R33='https://sswgw.org.cn/wcm.files/upload/CMSsszxw/202104/202104261055040.jpg'
}
foreach($id in $extra.Keys){
 $j=Get-Content -LiteralPath (Join-Path $assetRoot ($id+'.json')) -Raw|ConvertFrom-Json
 $file=Join-Path $assetRoot ($id+'.jpg')
 Invoke-WebRequest -Uri $extra[$id] -OutFile $file -TimeoutSec 25
 $j|Add-Member imagePath $file -Force
 $j|Add-Member imageUrl $extra[$id] -Force
 [IO.File]::WriteAllText((Join-Path $assetRoot ($id+'.json')),($j|ConvertTo-Json -Depth 5))
}
$j=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'sources.json') -Raw|ConvertFrom-Json|Where-Object id -eq 'R10'
$p=Invoke-WebRequest -Uri $j.url -TimeoutSec 25
[IO.File]::WriteAllText((Join-Path $assetRoot 'R10.html'),$p.Content)
$img=$p.Images|Where-Object {$_.alt -and $_.alt -ne '故宫博物院' -and $_.src -match '/Uploads/(Picture|image)/'}|Select-Object -First 1
$file=Join-Path $assetRoot 'R10.jpg'
Invoke-WebRequest -Uri $img.src -OutFile $file -TimeoutSec 25
$j|Add-Member imagePath $file
$j|Add-Member imageUrl $img.src
$j|Add-Member text ([Net.WebUtility]::HtmlDecode(($p.Content -replace '(?is)<script\b[^>]*>.*?</script>','' -replace '<[^>]+>',' ' -replace '\s+',' ')))
[IO.File]::WriteAllText((Join-Path $assetRoot 'R10.json'),($j|ConvertTo-Json -Depth 5))
$p=Invoke-WebRequest -Uri 'https://www.shanghaimuseum.net/mu/frontend/pg/article/aid/CI00000395' -TimeoutSec 25
[IO.File]::WriteAllText((Join-Path $assetRoot 'R06-frame.html'),$p.Content)
$p.Images|Select-Object src
