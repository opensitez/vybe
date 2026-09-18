#!/usr/bin/env bash
# vybe-test: bash/here_string_syntax/here_string_with_multiline_string
# A multi-line variable supplied via here-string delivers multiple newline-terminated lines to stdin.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
multiline_var="first
second
third"
count=0
while read -r line; do
    count=$(( count + 1 ))
done <<< "$multiline_var"
[ "$count" -eq 3 ] || fail "multiline here-string line count: want 3, got $count"
echo PASS
exit 0
