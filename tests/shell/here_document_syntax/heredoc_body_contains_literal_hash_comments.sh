#!/usr/bin/env bash
# vybe-test: bash/here_document_syntax/heredoc_body_contains_literal_hash_comments
# Lines starting with '#' inside a here-document are literal text, not comments, and are never stripped.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
doc=$(cat <<EOF
# Configuration Header
key = value
# Footer Note
EOF
)
expected="# Configuration Header"$'\n'"key = value"$'\n'"# Footer Note"
[ "$doc" = "$expected" ] || fail "heredoc hash lines stripped: got [$doc]"
echo PASS
exit 0
