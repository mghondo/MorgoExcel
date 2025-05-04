#!/bin/bash

# List of directories
DIRS=(
    "BUILD-IN"
    "DUTCHIE-IN"
    "DUTCHIE-OUT"
    "MORNING-IN"
    "MORNING-OUT"
    "MORNINGCOMPLETE"
    "MORNINGDROP"
    "ORDER-IN"
    "ORDER-OUT"
    "WEEKLY-IN"
    "WEEKLY-OUT"
    "WEEKLYCOMPLETE"
    "WEEKLYDROP"
)

# Loop through each directory and delete files except 'temp'
for dir in "${DIRS[@]}"; do
    if [ -d "$dir" ]; then
        find "$dir" -type f ! -name 'temp' -delete
    fi
done