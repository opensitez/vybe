#!/usr/bin/env bash
# vybe-test: bash/parameter_case_modification/caret_uppercases_first_or_all
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x="hello world"
[ "${x^}" = "Hello world" ] || fail "^: got [${x^}]"
[ "${x^^}" = "HELLO WORLD" ] || fail "^^: got [${x^^}]"
echo PASS
exit 0
