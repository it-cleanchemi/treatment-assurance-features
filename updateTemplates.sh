#!/bin/bash

# Read in Script IDs
# ScriptIDList.txt format:
#   Template Name:ScriptID:
#   ":" is the delimiter
#   Add ":" after ID to prevent newline issues
readarray -t lines < ScriptIDList.txt

# Loop through each array item to push updates
for line in "${lines[@]}"; do
    IFS=':'; arrIN=($line); unset IFS;
    echo ${arrIN[0]}
    
    # Make .clasp.json file
    echo "{" > .clasp.json
    echo '  "scriptId": "'${arrIN[1]}'",' >> .clasp.json
    echo '  "rootDir": "GAS"' >> .clasp.json
    echo "}" >> .clasp.json

    # Push updates to GAS
    clasp push
done

