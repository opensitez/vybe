#!/usr/bin/env bash
# vybe-test: bash/here_string_syntax/here_string_single_quoted_prevents_expansion
# Enclosing the here-string argument in single quotes prevents parameter and command expansion.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
var="unwanted"
read -r result <<< '$var and `pwd`'
[ "$result" = '$var and `pwd`' ] || fail "single quotes did not preserve literal characters: got [$result]"
echo PASS
exit 0
