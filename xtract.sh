#!/bin/bash
set -x
declare -a size=(`identify -ping -format "%w %h" ${1}.pdf[${2}]`)
footer=$(expr ${size[1]} - 140)
convert -gravity North -crop ${size[0]}x${footer}+0+0 ${1}.pdf[${2}] ${1}-${2}.png
