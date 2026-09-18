#!/usr/bin/env bash
# vybe-test: bash/builtin_mapfile_callbacks/optional_fragment
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=4
if ! command -v mapfile >/dev/null 2>&1; then
  exit 0
fi
map_count=0
map_callback(){ (( map_count += 1 )); }
mapfile -t -C map_callback -c 2 map_rows <<< 'alpha
beta
gamma
delta'
(( map_count >= 1 )) || fail "mapfile callback not invoked"
(( ${#map_rows[@]} >= 1 )) || fail "mapfile should populate output"
echo PASS
exit 0
