#!/usr/bin/env bash
# vybe-test: bash/parameter_pattern_replacement/newline_in_pattern_via_ansi_c_quoting
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=$'l1\nl2\nl3'
[ "${x//$'\n'/,}" = 'l1,l2,l3' ] || fail "got [${x//$'\n'/,}]"
echo PASS
exit 0
