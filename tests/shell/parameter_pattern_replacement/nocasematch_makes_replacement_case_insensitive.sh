#!/usr/bin/env bash
# vybe-test: bash/parameter_pattern_replacement/nocasematch_makes_replacement_case_insensitive
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=ABC
[ "${x/b/x}" = ABC ] || fail "case sensitive by default: got [${x/b/x}]"
shopt -s nocasematch
[ "${x/b/x}" = AxC ] || fail "nocasematch: got [${x/b/x}]"
echo PASS
exit 0
