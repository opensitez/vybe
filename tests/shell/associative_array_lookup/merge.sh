#!/usr/bin/env bash
# vybe-test: bash/associative_array_lookup/merge
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map
for ((i=0; i<14; i++)); do
  map["k$i"]="v$i"
done
idx=$((14 - 1))
[ "${map["k$idx"]}" = "v$idx" ] || fail "direct lookup failed"
unset map["k$((14 + 1))"]
[ -z "${map["k$((14 + 1))"]+x}" ] || fail "missing key unexpectedly set"
echo PASS
exit 0
