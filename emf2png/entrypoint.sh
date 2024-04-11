#!/bin/sh

set -o errexit

if [ -z "$1" ]
then
  INPUT_DIR="/github/workspace/test-cases"
else
  INPUT_DIR="$1"
fi

RUN_EMF2PNG=$(cat << EOM
  FILE=\$0
  OUTDIR=\$(dirname "\${FILE}")
  FILE_PNG="\${FILE%.*}.png"

  libreoffice --nologo --norestore --invisible --headless --convert-to png "\$FILE" --outdir "\$OUTDIR"

  convert -trim "\${FILE_PNG}" "\${FILE_PNG}"
EOM
)

# process all *.emf files from word-input and create index.md
find ${INPUT_DIR}/* -type f -name '*.emf' -exec sh -c "$RUN_EMF2PNG" {} ';'
