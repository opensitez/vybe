#!/usr/bin/env bash
# vybe-test: bash/bash_pattern_negation/case_match
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=13
shopt -s extglob
if (( IDX % 2 == 0 )); then
  [[ "readme.txt" == !(*.c|*.sh) ]] && fail "txt file should not satisfy negation"
  [[ "script.sh" == !(*.c|*.sh) ]] && fail "shell extension should be excluded"
else
  [[ "notes.md" == !(*.c|*.sh) ]] || fail "non-c files should match negation"
fi
shopt -u extglob
echo PASS
exit 0
