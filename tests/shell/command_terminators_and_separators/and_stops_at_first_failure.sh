#!/usr/bin/env bash
# vybe-test: bash/command_terminators_and_separators/and_stops_at_first_failure
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(true && echo a && false && echo b && echo c)
st=$?
[ "$out" = a ] || fail "want [a] got [$out]"
[ "$st" = 1 ] || fail "status of the list is the failing command's, want 1 got $st"
echo PASS
exit 0
