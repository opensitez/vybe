#!/usr/bin/env bash
# vybe-test: bash/parameter_pattern_replacement/extglob_and_character_classes_in_pattern
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=$'a b\tc'
[ "${x//[[:space:]]/_}" = 'a_b_c' ] || fail "class: got [${x//[[:space:]]/_}]"
shopt -s extglob
y=12ab345
[ "${y//+([0-9])/N}" = NabN ] || fail "extglob: got [${y//+([0-9])/N}]"
echo PASS
exit 0
