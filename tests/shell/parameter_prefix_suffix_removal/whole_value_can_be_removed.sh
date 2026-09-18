#!/usr/bin/env bash
# vybe-test: bash/parameter_prefix_suffix_removal/whole_value_can_be_removed
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=abc
[ -z "${x#abc}" ] || fail "literal whole match: got [${x#abc}]"
[ -z "${x%"$x"}" ] || fail "quoted self: got [${x%"$x"}]"
[ -z "${x##*}" ] || fail "longest * removes everything: got [${x##*}]"
[ "${x#*}" = abc ] || fail "shortest * matches the empty string: got [${x#*}]"
echo PASS
exit 0
