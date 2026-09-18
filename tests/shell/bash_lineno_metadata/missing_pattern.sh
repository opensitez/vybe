#!/usr/bin/env bash
# vybe-test: bash/bash_lineno_metadata/missing_pattern
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=5
lineno_capture=0
lineno_probe(){ lineno_capture=${BASH_LINENO[0]}; }
lineno_probe
(( lineno_capture > 0 )) || fail "BASH_LINENO must report caller line in function"
if (( IDX % 2 == 0 )); then
  if (( ${#BASH_LINENO[@]} < 1 )); then
    fail "BASH_LINENO metadata absent"
  fi
fi
echo PASS
exit 0
