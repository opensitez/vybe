#!/usr/bin/env bash
# vybe-test: bash/command_terminators_and_separators/semicolon_terminator_required_before_fi
# When 'fi' appears on the same line as the body command of an if branch, a semicolon terminator is required.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval 'if true; then echo 1 fi' 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "if statement without terminator before fi should fail parsing"

res=0
if true; then res=1; fi
[ "$res" -eq 1 ] || fail "if statement with semicolon before fi: want 1, got $res"
echo PASS
exit 0
