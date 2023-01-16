import csv
import codecs

media_url = 'https://www.galeria-anatomica.com/wp-content/uploads/product_images/'
din_formats = ['A0', 'A1', 'A2', 'A3', 'A4', 'A5']
prices = ['64,99', '44,99', '34,99', '24,99', '19,99', '13,99']
style_list = ['weiß', 'weiß mit Schatten', 'grau','grau mit Schatten', 'schwarz','vintage', 'vintage mit Schatten', 'schwarz-weiß', 'negativ']

def main():
    master_id = 30000

    product_dict = csv.DictReader(codecs.open('product-catalog.csv', 'r', 'utf-8'), delimiter='|', quotechar = "'") 
    product_list = list(product_dict)
    number_of_products = len(product_list)
    number_of_splits = 30
    records_per_split = int((number_of_products - (number_of_products % number_of_splits)) / number_of_splits)

    for i in range(1, number_of_splits + 1):
        product_import_file = codecs.open('woocommerce-import-' + str(i).zfill(2) + '.csv', 'w', 'utf-8')
        write_header(product_import_file)
        for product in product_list[(i - 1) * records_per_split:records_per_split * i]:
            write_product_master(product, master_id, product_import_file)
            write_product_variation(product, master_id, product_import_file)
            master_id = master_id + len(din_formats) * len(product['styles'].split(";")) + 1
        if i == number_of_splits and number_of_products % number_of_splits != 0:
            for product in product_list[(records_per_split * i):]:
                write_product_master(product, master_id, product_import_file)
                write_product_variation(product, master_id, product_import_file)
                master_id = master_id + len(din_formats) * len(product['styles'].split(";")) + 1
            
        product_import_file.close()

def write_header(output_file):
    header = "ID,Typ,Artikelnummer,Name,Veröffentlicht,Ist hervorgehoben?,Sichtbarkeit im Katalog,Kurzbeschreibung,Beschreibung," + '"' + "Datum, an dem Angebotspreis beginnt"  + '",' + '"' + "Datum, an dem Angebotspreis endet" + '"' + ",Steuerstatus,Steuerklasse,Vorrätig?,Lager,Geringe Lagermenge,Lieferrückstande erlaubt?,Nur einzeln verkaufen?,Gewicht (kg),Länge (cm),Breite (cm),Höhe (cm),Kundenbewertungen erlauben?,Hinweis zum Kauf,Angebotspreis,Regulärer Preis,Kategorien,Schlagwörter,Versandklasse,Bilder,Downloadlimit,Ablauftage des Downloads,Übergeordnetes Produkt,Gruppierte Produkte,Zusatzverkäufe,Cross-Sells (Querverkäufe),Externe URL,Button-Text,Position,Ist Dienstleistung?,Ist differenzbesteuert?,Versand kostenlos?,Regulärer Grundpreis,Angebotsgrundpreis,Grundpreis automatisch berechnen?,Einheit,Grundpreiseinheit,Produkteinheit,Warenkorbkurzbeschreibung,Lieferzeit,Streichpreis Hinweis,Angebotspreis Hinweis,Attribut 1 Name,Attribut 1 Wert(e),Attribut 1 Sichtbar,Attribut 1 Global,Attribut 1 Standard,Attribut 2 Name,Attribut 2 Wert(e),Attribut 2 Sichtbar,Attribut 2 Global,Attribut 2 Standard,Meta: _unit,Meta: _unit_base,Meta: _unit_product,Meta: _unit_price_auto,Meta: _unit_price_regular,Meta: _unit_price,Meta: _unit_price_sale,Meta: _sale_price_label,Meta: _sale_price_regular_label,Meta: _mini_desc,Meta: _min_age,Meta: _free_shipping,Meta: _service,Meta: _differential_taxation,Meta: _ts_gtin,Meta: _ts_mpn,Meta: _hs_code,Meta: _manufacture_country"
    output_file.write(header)
    output_file.write('\n')

def write_product_master(product, id, output_file):
    # load description for current product from file
    description = ''
    d = codecs.open(product['description'], 'r', 'utf-8')
    for line in d:
        description = description + ' ' + line.rstrip()
    description = description.lstrip()
    d.close()
    
    # generate quoted string with comma-separated list of images
    image_array = ''
    style_array = ''
    for ps in product['styles'].split(";"):
        if ps in style_list:
            index =  style_list.index(ps)
            image_array = image_array + media_url + product['id'] + '-PV-' + str(index + 1).zfill(2) + '.jpg' + ', ' 
            style_array = style_array + ps + ', '
    image_array = '"' + image_array[:-2] + '"'
    style_array = '"' + style_array[:-2] + '"'

    # generate quoted string with comma-separated list of formats
    format_array = ''
    for f in din_formats:
        format_array = format_array + f + ', '
    format_array = '"' + format_array[:-2] + '"'

    # fill dict with master details
    product_master = {
        'id': str(id),
        'typ': 'variable',
        'artikelnummer': product['id'],
        'name': product['name'],
        'veröffentlicht': '1',
        'ist_hervorgehoben': '0',
        'sichtbarkeit_im_katalog': 'visible',
        'kurzbeschreibung': '"' + product['short_description'] + '"',
        'beschreibung': '"' + description + '"',
        'angebotspreis_beginnt_am': '',
        'angebotspreis_endet_am': '',
        'steuerstatus': 'taxable',
        'steuerklasse': '',
        'vorrätig': '1',
        'lager': '',
        'geringe_lagermenge': '',
        'lieferrückstande_erlaubt': '0',
        'nur_einzeln_verkaufen': '0',
        'gewicht': '',
        'länge': '',
        'breite': '',
        'höhe': '',
        'kundenbewertungen_erlauben': '',
        'hinweis_zum_kauf': '',
        'angebotspreis': '',
        'regulärer_preis': '',
        'kategorien': product['categories'],
        'schlagwörter': '"' + product['keywords'] + '"',
        'versandklasse': '',
        'bilder': image_array,
        'downloadlimit': '',
        'ablauftage_des_downloads': '',
        'übergeordnetes_produkt': '',
        'gruppierte_produkte': '',
        'zusatzverkäufe': '',
        'cross-sells': '',
        'externe_url': '',
        'button_text': '',
        'position': '0',
        'ist_dienstleistung': '0',
        'ist_differenzbesteuert': '0',
        'versand_kostenlos': '0',
        'regulärer_grundpreis': '',
        'angebotsgrundpreis': '',
        'grundpreis_automatisch_berechnen': '0',
        'einheit': '',
        'grundpreiseinheit': '',
        'produkteinheit': '',
        'warenkorbkurzbeschreibung': '',
        'lieferzeit': '',
        'streichpreis_hinweis': 'Alter Preis:',
        'angebotspreis_hinweis': 'Neuer Preis:',
        'attribut_1_name': 'Format',
        'attribut_1_werte': format_array,
        'attribut_1_sichtbar': '1',
        'attribut_1_global': '0',
        'attribut_1_standard': 'A2',
        'attribut_2_name': 'Stil',
        'attribut_2_werte': style_array,
        'attribut_2_sichtbar': '1',
        'attribut_2_global': '0',
        'attribut_2_standard': '"' + style_array.split(",")[1].strip() + '"',
        'meta_unit': '',
        'meta_unit_base': '',
        'meta_unit_product': '',
        'meta_unit_price_auto': '0',
        'meta_unit_price_regular': '',
        'meta_unit_price': '',
        'meta_unit_price_sale': '',
        'meta_sale_price_label': 'old-price',
        'meta_sale_price_regular_label': 'new-price',
        'meta_mini_desc': '',
        'meta_min_age': '',
        'meta_free_shipping': 'no',
        'meta_service': 'no',
        'meta_differential_taxation': 'no',
        'meta_ts_gtin': '',
        'meta_ts_mpn': '',
        'meta_hs_code': '',
        'meta_manufacture_country': 'DE' 
    }
    output_file.write(",".join(product_master.values()))
    output_file.write('\n')

def write_product_variation(product, id, output_file):
    i = 0
    for ps in product['styles'].split(";"):
        if ps in style_list:
            index =  style_list.index(ps)
            for df in din_formats:
                i = i + 1
                # fill dict with background variation details
                product_variation = {
                    'id': str(id + i),
                    'typ': 'variation',
                    'artikelnummer': str(product['id']).rstrip() + '-' + df + '-' + str(index + 1).zfill(2),
                    'name': product['name'] + '-' + df + '-' + ps ,
                    'veröffentlicht': '1',
                    'ist_hervorgehoben': '0',
                    'sichtbarkeit_im_katalog': 'visible',
                    'kurzbeschreibung': '',
                    'beschreibung': '"' + product['short_description'] + '"',
                    'angebotspreis_beginnt_am': '',
                    'angebotspreis_endet_am': '',
                    'steuerstatus': 'taxable',
                    'steuerklasse': 'parent',
                    'vorrätig': '1',
                    'lager': '',
                    'geringe_lagermenge': '',
                    'lieferrückstande_erlaubt': '0',
                    'nur_einzeln_verkaufen': '0',
                    'gewicht': '0.2',
                    'länge': '',
                    'breite': '',
                    'höhe': '',
                    'kundenbewertungen_erlauben': '1',
                    'hinweis_zum_kauf': '',
                    'angebotspreis': '',
                    'regulärer_preis': '"' + prices[din_formats.index(df)] + '"',
                    'kategorien': '',
                    'schlagwörter': '',
                    'versandklasse': '',
                    'bilder': media_url + product['id'] + '-PV-' + str(index + 1).zfill(2) + '.jpg',
                    'downloadlimit': '',
                    'ablauftage_des_downloads': '',
                    'übergeordnetes_produkt': product['id'],
                    'gruppierte_produkte': '',
                    'zusatzverkäufe': '',
                    'cross-sells': '',
                    'externe_url': '',
                    'button_text': '',
                    'position': str(len(din_formats) * len(product['styles'].split(";")) - i + 1),
                    'ist_dienstleistung': '0',
                    'ist_differenzbesteuert': '0',
                    'versand_kostenlos': '0',
                    'regulärer_grundpreis':  '',
                    'angebotsgrundpreis': '',
                    'grundpreis_automatisch_berechnen': '',
                    'einheit': '0',
                    'grundpreiseinheit': '',
                    'produkteinheit': '',
                    'warenkorbkurzbeschreibung': product['name'] + '-' + df + '-' + ps,
                    'lieferzeit': '',
                    'streichpreis_hinweis': '',
                    'angebotspreis_hinweis': '',
                    'attribut_1_name': 'Format',
                    'attribut_1_werte': df,
                    'attribut_1_sichtbar': '',
                    'attribut_1_global': '0',
                    'attribut_1_standard': '',
                    'attribut_2_name': 'Stil',
                    'attribut_2_werte': '"' + ps + '"',
                    'attribut_2_sichtbar': '',
                    'attribut_2_global': '0',
                    'attribut_2_standard': '',
                    'meta_unit': '',
                    'meta_unit_base': '',
                    'meta_unit_product': '',
                    'meta_unit_price_auto': '1',
                    'meta_unit_price_regular': '',
                    'meta_unit_price': '',
                    'meta_unit_price_sale': '',
                    'meta_sale_price_label': '',
                    'meta_sale_price_regular_label': '',
                    'meta_mini_desc': str(product['id']).rstrip() + '-' + df + '-' + str(index + 1).zfill(2),
                    'meta_min_age': '',
                    'meta_free_shipping': '',
                    'meta_service': 'no',
                    'meta_differential_taxation': '',
                    'meta_ts_gtin': '',
                    'meta_ts_mpn': '',
                    'meta_hs_code': '',
                    'meta_manufacture_country': 'DE'
                }
                output_file.write(",".join(product_variation.values()))
                output_file.write('\n')

if __name__ == "__main__":
    main()

