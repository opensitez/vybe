#!/usr/bin/env bash
# vybe-test: bash/associative_array_quoting/child
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map
map["key with space 11"]="value with space"
map["key:semicolon11"]="value:semicolon"
[ "${map["key with space 11"]}" = "value with space" ] || fail "space key failed"
[ "${map["key:semicolon11"]}" = "value:semicolon" ] || fail "punctuation key failed"
echo PASS
exit 0
