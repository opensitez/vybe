#!/usr/bin/env bash
# vybe-test: bash/associative_array_quoting/subshell
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map
map["key with space 3"]="value with space"
map["key:semicolon3"]="value:semicolon"
[ "${map["key with space 3"]}" = "value with space" ] || fail "space key failed"
[ "${map["key:semicolon3"]}" = "value:semicolon" ] || fail "punctuation key failed"
echo PASS
exit 0
