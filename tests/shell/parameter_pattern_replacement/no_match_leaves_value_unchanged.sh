#!/usr/bin/env bash
# vybe-test: bash/parameter_pattern_replacement/no_match_leaves_value_unchanged
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=abc; e=
[ "${x/z/y}" = abc ] || fail "got [${x/z/y}]"
[ "${x//z}" = abc ] || fail "got [${x//z}]"
[ -z "${e/a/b}" ] || fail "empty value: got [${e/a/b}]"
echo PASS
exit 0
