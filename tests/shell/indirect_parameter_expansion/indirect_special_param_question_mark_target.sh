#!/usr/bin/env bash
# vybe-test: bash/indirect_parameter_expansion/indirect_special_param_question_mark_target
# The ${!ptr} expansion where ptr="?" retrieves the exit status of the last foreground command.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
ptr="?"
( exit 42 )
val="${!ptr}"
[ "$val" -eq 42 ] || fail "indirect \$? resolution failed: want 42, got [$val]"
echo PASS
exit 0
