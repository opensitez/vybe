#!/usr/bin/env bash
# vybe-test: bash/associative_array_lookup/chain
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map
for ((i=0; i<5; i++)); do
  map["k$i"]="v$i"
done
idx=$((5 - 1))
[ "${map["k$idx"]}" = "v$idx" ] || fail "direct lookup failed"
unset map["k$((5 + 1))"]
[ -z "${map["k$((5 + 1))"]+x}" ] || fail "missing key unexpectedly set"
echo PASS
exit 0
