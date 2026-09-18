#!/usr/bin/env bash
# vybe-test: bash/quoting_and_quote_removal/quotes_inside_command_substitution_are_independent
# Double quotes inside $(…) do not close the surrounding double quotes.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
out="$(echo "a  b")"
[ "$out" = "a  b" ] || fail "want [a  b] got [$out]"
n=$(count "$(echo "a b") c")
[ "$n" = 1 ] || fail "want 1 arg got $n"
echo PASS
exit 0
