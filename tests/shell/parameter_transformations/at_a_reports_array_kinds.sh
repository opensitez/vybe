#!/usr/bin/env bash
# vybe-test: bash/parameter_transformations/at_a_reports_array_kinds
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -a idx=(1)
declare -A map=([k]=v)
declare -ai nums=(1)
[ "${idx@a}" = a ] || fail "indexed: got [${idx@a}]"
[ "${map@a}" = A ] || fail "associative: got [${map@a}]"
[ "${nums@a}" = ai ] || fail "combined flags: got [${nums@a}]"
plain=x
[ -z "${plain@a}" ] || fail "no attributes: got [${plain@a}]"
echo PASS
exit 0
