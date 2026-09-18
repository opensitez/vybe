#!/usr/bin/env bash
# vybe-test: bash/truth_status_and_empty_values/true_and_false_builtins_via_variable
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
flag=false
if $flag; then r=y; else r=n; fi
[ "$r" = n ] || fail "false builtin via variable"
flag=true
if $flag; then r=y; else r=n; fi
[ "$r" = y ] || fail "true builtin via variable"
if "$flag"; then r=y; else r=n; fi
[ "$r" = y ] || fail "quoted variable also runs the builtin"
echo PASS
exit 0
