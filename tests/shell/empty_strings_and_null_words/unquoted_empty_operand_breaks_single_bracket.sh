#!/usr/bin/env bash
# vybe-test: bash/empty_strings_and_null_words/unquoted_empty_operand_breaks_single_bracket
# [ $x = "" ] with empty x becomes [ = ] and is an error; quoting fixes it and
# [[ never has the problem.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=
msg=$( [ $x = "" ] 2>&1 ); st=$?
[ "$st" -eq 2 ] || fail "want status 2 got $st"
[[ $msg == *"unary operator expected"* ]] || fail "got [$msg]"
[ "$x" = "" ] || fail "quoted comparison must be true"
[[ $x == "" ]] || fail "[[ handles the empty operand"
echo PASS
exit 0
