#!/usr/bin/env bash
# vybe-test: bash/parameter_transformations/at_Q_uses_ansi_c_form_for_control_characters
# A value with a tab or newline is quoted as $'...' so it can be reused as input.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=$'a\tb'
[ "${x@Q}" = "\$'a\\tb'" ] || fail "got [${x@Q}]"
eval "y=${x@Q}"
[ "$y" = "$x" ] || fail "round trip through eval failed"
e=
[ "${e@Q}" = "''" ] || fail "empty value: got [${e@Q}]"
echo PASS
exit 0
