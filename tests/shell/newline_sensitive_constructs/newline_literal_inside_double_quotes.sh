#!/usr/bin/env bash
# vybe-test: bash/newline_sensitive_constructs/newline_literal_inside_double_quotes
# An unescaped newline enclosed within double quotes is preserved literally while permitting expansions.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
var="middle"
s="prefix
$var
suffix"
expected="prefix"$'\n'"middle"$'\n'"suffix"
[ "$s" = "$expected" ] || fail "double quote literal newline failed: got [$s]"
echo PASS
exit 0
