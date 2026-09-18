#!/usr/bin/env bash
# vybe-test: bash/associative_array_lookup/disable
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map
for ((i=0; i<8; i++)); do
  map["k$i"]="v$i"
done
idx=$((8 - 1))
[ "${map["k$idx"]}" = "v$idx" ] || fail "direct lookup failed"
unset map["k$((8 + 1))"]
[ -z "${map["k$((8 + 1))"]+x}" ] || fail "missing key unexpectedly set"
echo PASS
exit 0
