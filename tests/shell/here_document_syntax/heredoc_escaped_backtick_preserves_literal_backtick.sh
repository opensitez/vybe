#!/usr/bin/env bash
# vybe-test: bash/here_document_syntax/heredoc_escaped_backtick_preserves_literal_backtick
# Preceding a backtick with a backslash in an unquoted heredoc outputs a literal backtick.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
doc=$(cat <<EOF
\`not command substitution\`
EOF
)
expected='`not command substitution`'
[ "$doc" = "$expected" ] || fail "escaped backtick failed: got [$doc]"
echo PASS
exit 0
