#!/usr/bin/env bash
# vybe-test: bash/parameter_expansion_basics/tilde_and_brace_from_expansion_stay_literal
# Brace and tilde expansion run before parameter expansion, so a ~ or {a,b}
# that comes from a variable's value is never expanded.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
t='~'
[ "$(echo $t)" = '~' ] || fail "tilde from value expanded: got [$(echo $t)]"
b='{a,b}'
[ "$(echo $b)" = '{a,b}' ] || fail "brace from value expanded: got [$(echo $b)]"
echo PASS
exit 0
