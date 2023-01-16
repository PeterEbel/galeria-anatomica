#!/bin/bash

for f1 in $1/*; do
   sed -i 's/http:\/\/192.168.178.65:9980/http:\/\/192.168.178.61:9980/g' "$f1"
done


