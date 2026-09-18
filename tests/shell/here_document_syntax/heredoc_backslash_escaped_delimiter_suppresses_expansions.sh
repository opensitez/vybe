#!/usr/bin/env bash
# vybe-test: bash/here_document_syntax/heredoc_backslash_escaped_delimiter_suppresses_expansions
# Prepending a backslash to the here-document delimiter suppresses all body expansions.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x="test"
body=$(cat <<\EOF
$x $(( 1 + 1 ))
EOF
)
expected='$x $(( 1 + 1 ))'
[ "$body" = "$expected" ] || fail "backslash-escaped delimiter did not suppress expansion: got [$body]"
echo PASS
exit 0
