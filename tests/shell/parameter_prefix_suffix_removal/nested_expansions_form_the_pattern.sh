#!/usr/bin/env bash
# vybe-test: bash/parameter_prefix_suffix_removal/nested_expansions_form_the_pattern
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=prefix-body.suffix
pre=prefix-
[ "${x#${pre}}" = body.suffix ] || fail "variable: got [${x#${pre}}]"
[ "${x%$(echo .suffix)}" = prefix-body ] || fail "command substitution: got [${x%$(echo .suffix)}]"
[ "${x#${x%%-*}-}" = body.suffix ] || fail "nested removal: got [${x#${x%%-*}-}]"
echo PASS
exit 0
