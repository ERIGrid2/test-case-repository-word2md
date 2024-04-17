#!/bin/sh

set -o errexit

if [ -z "$1" ]
then
  OUTPUT_DIR="/github/workspace/test-cases"
else
  OUTPUT_DIR="$1"
fi

mkdir -p ${OUTPUT_DIR}

python3 convert.py -r -e ./word-input ${OUTPUT_DIR}

find "${OUTPUT_DIR}"/* -type d -exec sh -c '
  DIRPATH=$0
  DIRNAME=$(basename "$DIRPATH")
  if [ ! -f "${DIRPATH}/index.md" ]
  then
    if [ ! -f "${DIRPATH}/_index.md" ]
    then
      echo Creating title link for directory: $DIRPATH with title: $DIRNAME
      cat > "${DIRPATH}/_index.md" <<EOF
---
title: "$DIRNAME"
linkTitle: "$DIRNAME"
weight: 5
---
EOF
    fi
  fi
' {} ${OUTPUT_DIR} ';'

RUN_EMF2PNG=$(cat << EOM
  FILE=\$0
  OUTDIR=\$(dirname "\${FILE}")
  FILE_PNG="\${FILE%.*}.png"
  libreoffice --nologo --norestore --invisible --headless --convert-to png "\$FILE" --outdir "\$OUTDIR"
  convert -trim "\${FILE_PNG}" "\${FILE_PNG}"
EOM
)

# process all *.emf files from word-input and create index.md
find ${OUTPUT_DIR}/* -type f -name '*.emf' -exec sh -c "$RUN_EMF2PNG" {} ';'
