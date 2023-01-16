import subprocess
import glob
import PIL
from wand.image import Image

BORDER = 400

# in_path = "D:\\Projekte\\Galeria-Anatomica\\Bilder\\Anatomie\\Tests\\"
# out_path = "D:\\Projekte\\Galeria-Anatomica\\Bilder\\Anatomie\\Tests\\Variations\\"
# background_path ="D:\\Projekte\\Galeria-Anatomica\\Bilder\\Background\\"

in_path = "D:\\Projekte\\Galeria-Anatomica\\Bilder\\Anatomie\\PNG\\Labeled\\"
out_path = "D:\\Projekte\\Galeria-Anatomica\\Bilder\\Anatomie\\Product\\Variations\\Labeled\\"
background_path ="D:\\Projekte\\Galeria-Anatomica\\Bilder\\Background\\"

portrait_formats =  [(1240,1748),(1748,2480),(2480,3508),(3505,4961),(4961,7016),(7016,9933),(9933,14043)]
landscape_formats = [(1748,1240),(2480,1748),(3508,2480),(4961,3505),(7016,4961),(9933,7016),(14043,9933)]

def find_format (_width, _height, _offset):
    i = 0
  # Portait
    if _width < _height:
        while _width + _offset > portrait_formats[i][0]: 
            i = i + 1
        while _height + _offset > portrait_formats[i][1]: 
            i = i + 1
        return portrait_formats[i][0], portrait_formats[i][1]
  # Landscape
    if _width >= _height:
        while _width + _offset > landscape_formats[i][0]: 
            i = i + 1
        while _height + _offset > landscape_formats[i][1]: 
            i = i + 1
        return landscape_formats[i][0], landscape_formats[i][1]

def process_files(_extension, _in_path, _out_path):
    for f in glob.glob(_in_path + _extension):
        in_file = f[-19:]
        basename = in_file[:-4]
        out_file = _out_path + basename + "-PV.png"
        out_file_variation_01 = _out_path + basename + "-PV-01.jpg"
        out_file_variation_02 = _out_path + basename + "-PV-02.jpg"
        out_file_variation_03 = _out_path + basename + "-PV-03.jpg"
        out_file_variation_04 = _out_path + basename + "-PV-04.jpg"
        out_file_variation_05 = _out_path + basename + "-PV-05.jpg"
        out_file_variation_06 = _out_path + basename + "-PV-06.jpg"
        out_file_variation_07 = _out_path + basename + "-PV-07.jpg"
        out_file_variation_08 = _out_path + basename + "-PV-08.jpg"
        out_file_variation_09 = _out_path + basename + "-PV-09.jpg"

        with Image(filename=f) as img:
            x, y = find_format(img.size[0], img.size[1], BORDER)
            print (basename + " " + str((img.size[0], img.size[1]))  + " " + str((x, y)))

          # Scale to allowed maximum width/height
            scale_factor = 1
          # Portrait
            if x < y:
                while img.size[0] * scale_factor < x - BORDER and img.size[1] * scale_factor < y - BORDER:
                    scale_factor = scale_factor + 0.1
          # Landscape
            if x >= y:
                while img.size[1] * scale_factor < y - BORDER and img.size[0] * scale_factor < x - BORDER:
                    scale_factor = scale_factor + 0.1
            scale_factor = scale_factor - 0.1    
            final_scale = scale_factor * 100    

            if x > y:
                background_file = background_path + "BG-VINTAGE-L-1240x874.jpg"
                resize = "1240x874"
            elif x <= y:
                background_file = background_path + "BG-VINTAGE-P-874x1240.jpg"
                resize = "874x1240"

          # Variation 01 - White
            parameters_01 = " -scale {3}% -background white -gravity center -extent {0}x{1} -resize {2} -quality 75% ".format(str(x), str(y), resize, final_scale)
            cmd = "convert {0} {1} {2}".format(f, parameters_01, out_file_variation_01) 
            subprocess.call(cmd, shell=True)
          # Variation 02 - White + Shadow
            parameters_02 = " -scale {3}% ( +clone -background black -shadow 40x20+40+30 -channel RGBA -blur 0x6 ) +swap ( +clone -background transparent -shadow 0x30-40-30 ) +swap -background transparent -layers merge +repage -background white -layers flatten -gravity center -extent {0}x{1} -resize {2} -quality 75%".format(str(x), str(y), resize, final_scale)
            cmd = "convert {0} {1} {2}".format(f, parameters_02, out_file_variation_02) 
            subprocess.call(cmd, shell=True)
          # Variation 03 - Grey
            parameters_03 = " -scale {3}% -background rgb(228,228,228) -gravity center -extent {0}x{1} -resize {2} -quality 75% ".format(str(x), str(y), resize, final_scale)
            cmd = "convert {0} {1} {2}".format(f, parameters_03, out_file_variation_03) 
            subprocess.call(cmd, shell=True)
          # Variation 04 - Grey + Shadow
            parameters_04 = " -scale {3}% ( +clone -background black -shadow 40x20+40+30 -channel RGBA -blur 0x6 ) +swap ( +clone -background transparent -shadow 0x30-40-30 ) +swap -background transparent -layers merge +repage -background rgb(228,228,228) -layers flatten -gravity center -extent {0}x{1} -resize {2} -quality 75%".format(str(x), str(y), resize, final_scale)
            cmd = "convert {0} {1} {2}".format(f, parameters_04, out_file_variation_04) 
            subprocess.call(cmd, shell=True)
          # Variation 05 - Black
            parameters_05 = " -scale {3}% -background black -gravity center -extent {0}x{1} -resize {2} -quality 75% ".format(str(x), str(y), resize, final_scale)
            cmd = "convert {0} {1} {2}".format(f, parameters_05, out_file_variation_05) 
            subprocess.call(cmd, shell=True)
          # Variation 06 - Vintage
            parameters_06 = " -scale {3}% -background transparent -gravity center -extent {0}x{1} -resize {2} -quality 75% ".format(str(x), str(y), resize, final_scale)
            cmd = "convert {0} {1} {2}".format(f, parameters_06, out_file) 
            subprocess.call(cmd, shell=True)
            cmd = "composite -gravity center {0} {1} {2}".format(out_file, background_file, out_file_variation_06) 
            subprocess.call(cmd, shell=True)
          # Variation 07 - Vintage + Shadow
            parameters_07 = " -scale {3}% ( +clone -background black -shadow 40x20+40+30 -channel RGBA -blur 0x6 ) +swap ( +clone -background transparent -shadow 0x30-40-30 ) +swap -background transparent -layers merge +repage -background transparent -layers flatten -gravity center -extent {0}x{1} -resize {2} -quality 75% ".format(str(x), str(y), resize, final_scale)
            cmd = "convert {0} {1} {2}".format(f, parameters_07, out_file) 
            subprocess.call(cmd, shell=True)
            cmd = "composite -gravity center {0} {1} {2}".format(out_file, background_file, out_file_variation_07) 
            subprocess.call(cmd, shell=True)
          # Variation 08 - Black and White
            parameters_08 = " -scale {3}% -background white -colorspace gray -gravity center -extent {0}x{1} -resize {2} -quality 75% ".format(str(x), str(y), resize, final_scale)
            cmd = "convert {0} {1} {2}".format(f, parameters_08, out_file_variation_08) 
            subprocess.call(cmd, shell=True)
          # Variation 09 - Negativ
            parameters_09 = " -scale {3}% -background black -colorspace gray -channel RGB -negate -gravity center -extent {0}x{1} -resize {2} -quality 75% ".format(str(x), str(y), resize, final_scale)
            cmd = "convert {0} {1} {2}".format(f, parameters_09, out_file_variation_09) 
            subprocess.call(cmd, shell=True)

def main():
    extension = "*.png"
    process_files(extension, in_path, out_path)

if __name__ == "__main__":
    main()


