#!/bin/sh

set -o errexit

if [ -z "$1" ]
then
  OUTPUT_DIR="/github/workspace/test-cases"
else
  OUTPUT_DIR="$1"
fi

mkdir -p ${OUTPUT_DIR}


# RUN_WORD2MD=$(cat << EOM
#   FILE=\$0
#   OUTPUT_DIR="\$1"
#   PREFIX="\$2"
#   DIRNAMEPREFIX=\$(dirname "\${FILE}")
#   IS_TLD=\$(echo \${DIRNAMEPREFIX} | grep "/")
#   if [ \$? -eq 1 ]; then
#     DIRNAME=/
#   else
#     DIRNAME=/\${DIRNAMEPREFIX#\$PREFIX}
#   fi
#   mkdir -p "\${OUTPUT_DIR}/\${DIRNAME}/"
#   OUTPUT_FILE_DIRNAME="\${OUTPUT_DIR}\${DIRNAME}"
#   echo ${OUTPUT_FILE_DIRNAME}
#   python3 convert.py "\$FILE" "\${OUTPUT_FILE_DIRNAME}"
# EOM
# )

# process all *.docx files from word-input and create index.md
# find word-input/* -type f -name '*.docx' -exec sh -c "$RUN_WORD2MD" {} ${OUTPUT_DIR} "word-input/" ';'

python3 convert.py -r -e ./word-input ${OUTPUT_DIR}

find "${OUTPUT_DIR}" -type d -exec sh -c '
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
