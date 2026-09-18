#!/usr/bin/env bash
# vybe-test: bash/word_splitting/no_splitting_inside_double_brackets_but_single_bracket_splits
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x='a b'
[[ $x == "a b" ]] || fail "[[ must not split \$x"
msg=$( [ $x = "a b" ] 2>&1 ); st=$?
[ "$st" -eq 2 ] || fail "[ splits and sees too many arguments, want status 2 got $st"
[[ $msg == *"too many arguments"* ]] || fail "got [$msg]"
echo PASS
exit 0
