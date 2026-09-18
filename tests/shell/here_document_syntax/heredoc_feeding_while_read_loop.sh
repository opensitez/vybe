#!/usr/bin/env bash
# vybe-test: bash/here_document_syntax/heredoc_feeding_while_read_loop
# A here-document can directly feed the standard input of a while read loop compound command.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
acc=""
while read -r item; do
    acc+="${item},"
done <<EOF
alpha
beta
gamma
EOF
[ "$acc" = "alpha,beta,gamma," ] || fail "while loop with heredoc failed: got [$acc]"
echo PASS
exit 0
