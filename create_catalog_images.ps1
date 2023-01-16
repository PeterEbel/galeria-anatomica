$catalog_in_files = Get-ChildItem "D:\Projekte\Galeria-Anatomica\Bilder\Anatomie\Labeled\" -Filter *.jpg
$catalog_out_path = "D:\Projekte\Galeria Anatomica\Bilder\Anatomie\Catalog\"

foreach ($f in $catalog_in_files) {
    $filename = $f.BaseName
    $size = @()
    $size = (identify -ping -format "%w %h" ($f.Fullname)).Split(" ")
    $width = [int] $size[0]
    $height = [int] $size[1]
    $max = [math]::max($width,$height)
    Write-Output ($filename + ".jpg")
#    convert ($f.FullName) -background white -gravity center -extent ([string] $max + "x" + [string] $max) -resize 256x256 -strip -quality 75% ($catalog_out_path + $filename.Substring(0,13)  + "C1.jpg")
    convert ($f.FullName) -gravity center -extent ([string] $max + "x" + [string] $max) -resize 256x256 -strip -quality 75% ($catalog_out_path + $filename  + "-WR.jpg")
    parameters_02 = " -scale {3}% ( +clone -background black -shadow 40x20+40+30 -channel RGBA -blur 0x6 ) +swap ( +clone -background transparent -shadow 0x30-40-30 ) +swap -background transparent -layers merge +repage -background white -layers flatten -gravity center -extent {0}x{1} -resize {2} -quality 75%".format(str(x), str(y), resize, final_scale)

}



