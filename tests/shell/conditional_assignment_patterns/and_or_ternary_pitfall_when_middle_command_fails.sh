#!/usr/bin/env bash
# vybe-test: bash/conditional_assignment_patterns/and_or_ternary_pitfall_when_middle_command_fails
# cond && a=1 || a=2 is safe only because plain assignments never fail; if the
# middle command can fail, the "else" part runs too.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
true && a=1 || a=2
[ "$a" -eq 1 ] || fail "plain assignment: got $a"
true && b=$(false) || b=2
[ "$b" -eq 2 ] || fail "failing middle: the || branch overwrote b, got [$b]"
echo PASS
exit 0
