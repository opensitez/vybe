#!/usr/bin/env bash
# vybe-test: bash/parameter_transformations/at_U_and_at_L_change_whole_string
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x="Hello World"
[ "${x@U}" = "HELLO WORLD" ] || fail "@U: got [${x@U}]"
[ "${x@L}" = "hello world" ] || fail "@L: got [${x@L}]"
echo PASS
exit 0
