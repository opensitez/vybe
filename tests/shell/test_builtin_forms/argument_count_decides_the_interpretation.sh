#!/usr/bin/env bash
# vybe-test: bash/test_builtin_forms/argument_count_decides_the_interpretation
# 0 args: false. 1 arg: true if non-empty (even "-f" or "!"). 2 args: unary
# operator or ! string. 3 args: binary. More: usually an error.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ ] && fail "no arguments must be false"
[ 0 ] && [ -f ] && [ ! ] || fail "one non-empty argument is true"
[ ! "" ] || fail "! empty is true"
[ ! a ] && fail "! non-empty is false"
[ a = a ] || fail "three arguments: binary"
[ a = b c ] 2>/dev/null; st=$?
[ "$st" -eq 2 ] || fail "four arguments: want status 2 got $st"
echo PASS
exit 0
