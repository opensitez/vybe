#!/usr/bin/env bash
# vybe-test: bash/here_document_syntax/heredoc_unquoted_delimiter_expands_parameters
# An unquoted here-document delimiter expands shell parameter expressions in its body.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
user="Alice"
role="Dev"
body=$(cat <<EOF
User: $user
Role: ${role}
EOF
)
expected="User: Alice"$'\n'"Role: Dev"
[ "$body" = "$expected" ] || fail "unquoted heredoc expansion failed: got [$body]"
echo PASS
exit 0
