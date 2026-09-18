#!/usr/bin/env bash
# vybe-test: bash/command_substitution/inherits_variables_and_positionals
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- p1 p2
v=unexported
f() { echo "fn:$1"; }
out=$(echo "$v $1 $# $(f x)")
[ "$out" = "unexported p1 2 fn:x" ] || fail "got [$out]"
echo PASS
exit 0
