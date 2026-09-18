#!/usr/bin/env bash
# vybe-test: bash/here_string_syntax/here_string_feeds_while_read_loop
# A multiline string supplied via <<< streams into a while read loop iteration.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
acc=""
while read -r entry; do
    acc+="${entry};"
done <<< $'row_1\nrow_2\nrow_3'
[ "$acc" = "row_1;row_2;row_3;" ] || fail "while loop over here-string failed: got [$acc]"
echo PASS
exit 0
