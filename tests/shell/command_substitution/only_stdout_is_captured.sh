#!/usr/bin/env bash
# vybe-test: bash/command_substitution/only_stdout_is_captured
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=$( { echo out; echo err >&2; } 2>/dev/null )
[ "$x" = out ] || fail "stderr must not be captured: got [$x]"
y=$( { echo out; echo err >&2; } 2>&1 )
[ "$y" = $'out\nerr' ] || fail "redirected stderr is captured: got [$y]"
echo PASS
exit 0
