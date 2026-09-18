#!/usr/bin/env bash
# vybe-test: bash/parameter_expansion_basics/expanded_word_can_become_command_and_arguments
# After expansion and word splitting the first resulting word is the command
# name; a variable may therefore hold a whole simple command.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
cmd="printf %s|"
out=$($cmd a b)
[ "$out" = 'a|b|' ] || fail "want [a|b|] got [$out]"
full="echo x   y"
out=$($full)
[ "$out" = 'x y' ] || fail "want [x y] got [$out]"
echo PASS
exit 0
