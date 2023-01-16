#!/bin/bash

for f1 in $1/*.docx; do
    f2="${f1%.docx}"
    mammoth "$f1" "$f2" --style-map=custom-style-map
    sed -i 's/#/"/g' "$f2"
    sed -i 's/+++/</g' "$f2"
    sed -i 's/---/>/g' "$f2"
    sed -i 's/&lt;/</g' "$f2"
    sed -i 's/&gt;/>/g' "$f2"
    sed -i 's/&quot;/"/g' "$f2"
    sed -i 's/<h2>[^P]*<\/h2>//g' "$f2"
    sed -i 's/\xC2\xA0/\&nbsp;/g' "$f2"
    if [ -n "$2" ]; then
        if [ "$2" = "public" ]; then
            sed -i 's/http:\/\/192.168.178.65:9980/https:\/\/www.galeria-anatomica.com/g' "$f2"
        fi
        if [ "$2" = "vm" ]; then
            sed -i 's/https:\/\/www.galeria-anatomica.com/http:\/\/192.168.178.65:9980/g' "$f2"
        fi
        if [ "$2" = "pi" ]; then
            sed -i 's/http:\/\/192.168.178.65:9980/http:\/\/192.168.178.61:9980/g' "$f2"
        fi        
        
    fi
done


