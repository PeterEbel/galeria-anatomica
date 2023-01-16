#pylint: disable=unused-variable

import io
import datetime
import os
import glob
import unicodedata
import string

import mysql.connector

# create database cursor
mydb = mysql.connector.connect(
    host="192.168.178.61",
    user="root",
    password="xxxxx",
    database="wp_galeria_anatomica"
)
mycursor = mydb.cursor()

valid_filename_chars = "-_.() %s%s" % (string.ascii_letters, string.digits)

def main():

    global mycursor

    # mycursor.execute("SELECT post_title FROM wp_posts WHERE post_type = 'yada_wiki'")
    # max_post_id = mycursor.fetchone()[0]
    # mycursor.execute("SELECT distinct(post_type) FROM wp_posts ORDER BY post_type")
    # post_types = mycursor.fetchall()
    # print(post_types)
    # rows_deleted = mycursor.execute("DELETE FROM wp_posts WHERE post_type = 'yada_wiki'")
    # mydb.commit()
    # print(mycursor.rowcount, " record(s) deleted")    

    # reset_database()
    upload_posts('/home/peter/Projects/galeria_anatomica/posts/public/*')
   

def upload_posts(path_to_posts):
    
  # Unordered list of files with posts to be batch uploaded
    tmp_files = glob.glob(path_to_posts)
    unsorted_file_dict = {}

  # Create dictionary with ordered keys and filenames
    for file in tmp_files:
        heading_number = os.path.basename(file).split(' - ')[0]
        heading_text = os.path.basename(file).split(' - ')[1]
        parent_padded = (heading_number.split('.')[0]).zfill(3)
        child_padded = (heading_number.split('.')[1]).zfill(3)
        new_heading_number = parent_padded + '.' + child_padded
        unsorted_file_dict[new_heading_number] = file
    
    sorted_file_dict = dict(sorted(unsorted_file_dict.items()))

    with io.open("batch_posts.sql", mode="w", encoding="utf-8") as f_out:

      # Generate parent and child IDs
        for key, file in sorted_file_dict.items():

            parent_id = int(key.split('.')[0]) * 1000
            child_id = int(key.split('.')[1]) - 1
            heading_text = os.path.basename(file).split(' - ')[1]

          # Get dates and times
            dt_now = datetime.datetime.now()
            dt_gmt = datetime.datetime.now() - datetime.timedelta(minutes=60)
            current_dt_now = dt_now.strftime("%Y-%m-%d %H:%M:%S")
            current_dt_gmt = dt_gmt.strftime("%Y-%m-%d %H:%M:%S")

          # Read the content of the file
            with io.open(file, mode="r", encoding="utf-8") as f:
                post_content = f.read()
                f.close()

            if heading_text in ['Willkommen bei Galeria Anatomica', 'Mitwirkende', 'Showrooms', 'Shop', 'Praxiseinrichtung', 'Impressum', 'Datenschutzerklärung', 'Widerrufsbelehrung', 'AGB', 'Spendenprojekt', 'Wer wir sind', 'FAQ']:
                child_id = 0
                post_type = 'page'
                guid_base = 'https://www.galeria-anatomica.com/?post_type=page;p='
                if heading_text == 'Willkommen bei Galeria Anatomica': parent_id = 1
                if heading_text == 'Impressum': parent_id = 2
                if heading_text == 'Datenschutzerklärung': parent_id = 3
                if heading_text == 'Widerrufsbelehrung': parent_id = 4
                if heading_text == 'Mitwirkende': parent_id = 5
                if heading_text == 'Showrooms': parent_id = 6
                if heading_text == 'Spendenprojekt': parent_id = 7
                if heading_text == 'Wer wir sind': parent_id = 8
                if heading_text == 'FAQ': parent_id = 9
                if heading_text == 'AGB': parent_id = 10
                if heading_text == 'Shop': parent_id = 11
                if heading_text == 'Praxiseinrichtung': parent_id = 12


            else:
                post_type = 'yada_wiki'    
                guid_base = 'https://www.galeria-anatomica.com/?post_type=yada_wiki;p='

            menu_order = parent_id + child_id
            post = {
                'id': menu_order,
                'post_author': 1,
                'post_date': current_dt_now,
                'post_date_gmt': current_dt_gmt,
                'post_content': post_content,
                'post_title': heading_text,
                'post_excerpt': '',
                'post_status': 'publish',
                'comment_status': 'closed',
                'ping_status': 'closed',
                'post_password': '',
                'post_name': sanitize_postname(heading_text),
                'to_ping': '',
                'pinged': '',
                'post_modified': current_dt_now,
                'post_modified_gmt': current_dt_gmt,
                'post_content_filtered': '',
                'post_parent': parent_id,
                'guid': guid_base + str(menu_order),
                'menu_order': menu_order,
                'post_type': post_type,
                'post_mime_type': '',
                'comment_count': 0
            }

            stmt = "INSERT INTO `{table}` ({columns}) VALUES {values};".format(table='wp_posts', columns=", ".join(post.keys()), values=tuple(post.values()))
            f_out.writelines(stmt + '\n')

            # placeholder = ", ".join(["%s"] * len(post))
            # stmt = "INSERT INTO `{table}` ({columns}) VALUES ({values});".format(table='wp_posts', columns=",".join(post.keys()), values=placeholder)
            # mycursor.execute(stmt, list(post.values()))
            # mydb.commit()

    f_out.close()    
            

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

def reset_database():

    global mycursor

    mycursor.execute("TRUNCATE TABLE wp_commentmeta")
    mydb.commit()
    mycursor.execute("TRUNCATE TABLE wp_comments")
    mydb.commit()
    mycursor.execute("TRUNCATE TABLE wp_postmeta")
    mydb.commit()
    mycursor.execute("TRUNCATE TABLE wp_posts")
    mydb.commit()
    mycursor.execute("TRUNCATE TABLE wp_termmeta")
    mydb.commit()
    mycursor.execute("TRUNCATE TABLE wp_term_relationships")
    mydb.commit()
    mycursor.execute("TRUNCATE TABLE wp_terms")
    mydb.commit()
    mycursor.execute("TRUNCATE TABLE wp_term_taxonomy")
    mydb.commit()

if __name__ == "__main__":
    main()

