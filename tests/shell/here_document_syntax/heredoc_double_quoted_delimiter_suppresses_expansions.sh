#!/usr/bin/env bash
# vybe-test: bash/here_document_syntax/heredoc_double_quoted_delimiter_suppresses_expansions
# Double-quoting the here-document delimiter suppresses parameter and command expansions in the body.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
num=42
body=$(cat <<"LIMIT"
Total: $num or $(printf 'calc')
LIMIT
)
expected='Total: $num or $(printf '"'calc')"
[ "$body" = "$expected" ] || fail "double-quoted delimiter did not suppress expansion: got [$body]"
echo PASS
exit 0
