#!/usr/bin/env bash
# vybe-test: bash/parameter_prefix_suffix_removal/removal_is_case_sensitive_even_with_nocasematch
# nocasematch affects case, [[ and pattern substitution, but not # and %.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=abc
[ "${x#A}" = abc ] || fail "got [${x#A}]"
shopt -s nocasematch
[ "${x#A}" = abc ] || fail "nocasematch must not apply: got [${x#A}]"
[ "${x%C}" = abc ] || fail "nocasematch must not apply: got [${x%C}]"
echo PASS
exit 0
