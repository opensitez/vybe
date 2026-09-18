#!/usr/bin/env bash
# vybe-test: bash/parameter_case_modification/tilde_toggles_case
# ${x~} toggles the first character, ${x~~} toggles every character.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=aBc
[ "${x~}" = ABc ] || fail "~: got [${x~}]"
[ "${x~~}" = AbC ] || fail "~~: got [${x~~}]"
echo PASS
exit 0
