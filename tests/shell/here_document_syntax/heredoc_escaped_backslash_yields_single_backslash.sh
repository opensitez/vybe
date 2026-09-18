#!/usr/bin/env bash
# vybe-test: bash/here_document_syntax/heredoc_escaped_backslash_yields_single_backslash
# A double backslash '\\' in an unquoted here-document outputs a single backslash.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
doc=$(cat <<EOF
path\\to\\dir
EOF
)
expected='path\to\dir'
[ "$doc" = "$expected" ] || fail "double backslash failed: got [$doc]"
echo PASS
exit 0
