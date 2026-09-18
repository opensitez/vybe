#!/usr/bin/env bash
# vybe-test: bash/builtin_mapfile_delimiters/case_match
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=13
if ! command -v mapfile >/dev/null 2>&1; then
  exit 0
fi
mapfile -t -d ':' map_cols <<< 'red:green:blue:yellow:'
(( ${#map_cols[@]} >= 3 )) || fail "delimiter mapfile should split into parts"
[[ ${map_cols[0]} == red ]] || fail "delimiter split first part"
[[ ${map_cols[1]} == green ]] || fail "delimiter split second part"
echo PASS
exit 0
