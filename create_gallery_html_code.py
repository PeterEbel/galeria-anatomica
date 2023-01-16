import csv
import codecs

media_url = 'http://192.168.178.61:9980/wp-content/uploads/product_images/'

def main():
    product_dict = csv.DictReader(codecs.open('product-catalog.csv', 'r', 'utf-8'), delimiter='|', quotechar = "'") 
    gallery_html_output_file = codecs.open('gallery.html', 'w', 'utf-8')

    write_start_tag(gallery_html_output_file)
    for product in product_dict:
        write_html(product, gallery_html_output_file)
    write_end_tag(gallery_html_output_file)
    gallery_html_output_file.close()

def write_start_tag(output_file):
    start_tag = '<div class=' + '"' + 'flexbox-container-gallery' + '"' + '>'
    output_file.write(start_tag)
    output_file.write('\n')

def write_end_tag(output_file):
    end_tag = '</div>'
    output_file.write(end_tag)
    output_file.write('\n')

def write_html(product, output_file):
    html = '<div><a href="http://192.168.178.61:9980/produkt/' +  sanitize_postname(product['short_description']) + '/' + '">' + \
           '<img class=' + '"' + 'product-gallery' + '"' + ' title=' + '"' + product['description'] + '"' + ' ' + \
           'src=' + '"' + media_url + product['id'] + '-WR.jpg' + '"' + ' ' + 'alt=' + '"' + product['description'] + '"' + '/></a></div>'
    # html = '    <div class=' + '"' + 'flexbox-container-toc' + '">' + '<a href="http://192.168.178.61:9980/produkt/' +  sanitize_postname(product['short_description']) + '/' + '">' + \
    #        '<img class=' + '"' + 'product-reference' + '"' + ' title=' + '"' + product['description'] + '"' + ' ' + \
    #        'src=' + '"' + media_url + product['id'] + '-PV-02.jpg' + '"' + ' ' + 'alt=' + '"' + product['description'] + '"' + '/></a></div>'
    output_file.write(html)
    output_file.write('\n')

def sanitize_postname(file):
    charDict = { 
        u'Ä': 'Ae',
        u'Ö': 'Oe',
        u'Ü': 'Ue',
        u'ä': 'ae',
        u'ö': 'oe',
        u'ü': 'ue',
        u'ß': 'ss',
        u' ': '-',
        u',':''
    }
    umap = {ord(key):val for key, val in charDict.items()}
    file = file.translate(umap)
    return file.lower()

if __name__ == "__main__":
    main()

