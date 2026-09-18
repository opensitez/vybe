#!/usr/bin/env bash
# vybe-test: bash/associative_array_lookup/override
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map
for ((i=0; i<4; i++)); do
  map["k$i"]="v$i"
done
idx=$((4 - 1))
[ "${map["k$idx"]}" = "v$idx" ] || fail "direct lookup failed"
unset map["k$((4 + 1))"]
[ -z "${map["k$((4 + 1))"]+x}" ] || fail "missing key unexpectedly set"
echo PASS
exit 0
