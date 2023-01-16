$catalog_in_files = Get-ChildItem "D:\Projekte\Galeria Anatomica\Bilder\Anatomie\Labeled\" -Filter *.jpg
$catalog_out_path = "D:\Projekte\Galeria Anatomica\Bilder\Anatomie\Product\Master\"

foreach ($f in $catalog_in_files) {
    $filename = $f.BaseName
    $size = @()
    $size = (identify -ping -format "%w %h" ($f.Fullname)).Split(" ")
    $width = [int] $size[0]
    $height = [int] $size[1]
    $max = [math]::max($width,$height)
    Write-Output ($filename + "-PM.jpg")
#    convert ($f.FullName) -background white -gravity center -extent ([string] $max + "x" + [string] $max) -resize 256x256 -strip -quality 75% ($catalog_out_path + $filename.Substring(0,13)  + "C1.jpg")
    convert ($f.FullName) -gravity center -extent ([string] $max + "x" + [string] $max) -resize 1024x1024 -strip -quality 75% ($catalog_out_path + $filename  + "-PM.jpg")
}



