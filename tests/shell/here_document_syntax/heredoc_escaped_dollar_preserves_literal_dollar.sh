#!/usr/bin/env bash
# vybe-test: bash/here_document_syntax/heredoc_escaped_dollar_preserves_literal_dollar
# Preceding a '$' with a backslash in an unquoted heredoc outputs a literal dollar sign without expansion.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
var="expanded"
doc=$(cat <<EOF
literal \$var is not $var
EOF
)
expected="literal \$var is not expanded"
[ "$doc" = "$expected" ] || fail "escaped dollar failed: got [$doc]"
echo PASS
exit 0
