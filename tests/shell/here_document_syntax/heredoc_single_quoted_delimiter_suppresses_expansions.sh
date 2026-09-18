#!/usr/bin/env bash
# vybe-test: bash/here_document_syntax/heredoc_single_quoted_delimiter_suppresses_expansions
# Single-quoting any part of the heredoc delimiter suppresses variable and command expansions.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
secret="hidden"
body=$(cat <<'EOF'
Value is $secret and `echo nested`
EOF
)
expected='Value is $secret and `echo nested`'
[ "$body" = "$expected" ] || fail "single-quoted heredoc delimiter did not suppress expansion: got [$body]"
echo PASS
exit 0
