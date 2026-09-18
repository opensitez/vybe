#!/usr/bin/env bash
# vybe-test: bash/quoting_and_quote_removal/newline_inside_quotes_is_part_of_word
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
x="a
b"
[ "${#x}" -eq 3 ] || fail "length want 3 got ${#x}"
n=$(count "$x")
[ "$n" = 1 ] || fail "quoted newline must not split, got $n args"
n=$(count $x)
[ "$n" = 2 ] || fail "unquoted newline is IFS whitespace, want 2 got $n"
echo PASS
exit 0
