#!/usr/bin/env bash
# vybe-test: bash/command_terminators_and_separators/newline_allowed_after_and_or_operator
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(true &&
echo a ||
echo b)
[ "$out" = a ] || fail "got [$out]"
eval 'true
&& echo c' 2>/dev/null; st=$?
[ "$st" -ne 0 ] || fail "operator at start of a line must be a syntax error"
echo PASS
exit 0
