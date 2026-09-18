#!/usr/bin/env bash
# vybe-test: bash/parameter_prefix_suffix_removal/empty_pattern_and_empty_value
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=abc; e=; unset u
[ "${x#}" = abc ] || fail "empty pattern removes nothing: got [${x#}]"
[ "${x%%}" = abc ] || fail "empty pattern removes nothing: got [${x%%}]"
[ -z "${e#a}" ] || fail "empty value: got [${e#a}]"
[ -z "${u%*}" ] || fail "unset value: got [${u%*}]"
echo PASS
exit 0
