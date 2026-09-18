#!/usr/bin/env bash
# vybe-test: bash/whitespace_and_newlines/newline_after_then_and_do_keywords
# Newlines after 'then' and 'do' keywords allow body commands without an explicit semicolon.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
if true
then
    x=10
fi
[ "$x" -eq 10 ] || fail "if-then with newlines: want 10, got $x"

count=0
for item in a b c
do
    count=$((count + 1))
done
[ "$count" -eq 3 ] || fail "for-do with newlines: want 3, got $count"
echo PASS
exit 0
