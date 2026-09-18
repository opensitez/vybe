#!/usr/bin/env bash
# vybe-test: bash/arithmetic_expansion/legacy_dollar_bracket_form_still_works
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ "$[1+1]" -eq 2 ] || fail "got [$[1+1]]"
x=4
[ "$[x*x]" -eq 16 ] || fail "got [$[x*x]]"
echo PASS
exit 0
