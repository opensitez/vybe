#!/usr/bin/env bash
# vybe-test: bash/quoting_and_quote_removal/quote_removal_in_associative_array_index
# Quotes used inside associative array index brackets are removed before key lookup.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map
map["my key"]="value1"
[ "${map['my key']}" = "value1" ] || fail "single quote lookup: got [${map['my key']}]"
[ "${map[my key]}" = "value1" ] || fail "unquoted space lookup: got [${map[my key]}]"
echo PASS
exit 0
