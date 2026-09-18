#!/usr/bin/env bash
# vybe-test: bash/command_terminators_and_separators/pipe_binds_tighter_than_and_or
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(false | true && echo yes)
[ "$out" = yes ] || fail "pipeline status is last command's, want [yes] got [$out]"
out=$(true | false || echo no)
[ "$out" = no ] || fail "want [no] got [$out]"
echo PASS
exit 0
