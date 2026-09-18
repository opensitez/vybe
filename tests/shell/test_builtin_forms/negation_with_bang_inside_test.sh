#!/usr/bin/env bash
# vybe-test: bash/test_builtin_forms/negation_with_bang_inside_test
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ ! -e /nonexistent_zz ] || fail "! -e"
[ ! a = b ] || fail "! binary"
[ ! a = a ] && fail "! true binary"
[ ! ! a = a ] || fail "double negation"
echo PASS
exit 0
