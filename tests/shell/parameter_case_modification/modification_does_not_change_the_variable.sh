#!/usr/bin/env bash
# vybe-test: bash/parameter_case_modification/modification_does_not_change_the_variable
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=abc
: "${x^^}"
[ "$x" = abc ] || fail "variable mutated to [$x]"
y=${x^^}
[ "$y" = ABC ] && [ "$x" = abc ] || fail "assignment of result must not touch source"
echo PASS
exit 0
