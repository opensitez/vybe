#!/usr/bin/env bash
# vybe-test: bash/parameter_case_modification/comma_lowercases_first_or_all
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x="HELLO World"
[ "${x,}" = "hELLO World" ] || fail ",: got [${x,}]"
[ "${x,,}" = "hello world" ] || fail ",,: got [${x,,}]"
echo PASS
exit 0
