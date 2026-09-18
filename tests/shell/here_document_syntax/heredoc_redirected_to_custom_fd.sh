#!/usr/bin/env bash
# vybe-test: bash/here_document_syntax/heredoc_redirected_to_custom_fd
# A here-document can be directed to a specific file descriptor like 3<<EOF and read with read -u 3.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
exec 3<<EOF
payload_on_fd_3
EOF
read -u 3 -r line
exec 3<&-
[ "$line" = "payload_on_fd_3" ] || fail "custom fd heredoc read failed: got [$line]"
echo PASS
exit 0
