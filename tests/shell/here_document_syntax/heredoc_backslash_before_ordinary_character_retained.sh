#!/usr/bin/env bash
# vybe-test: bash/here_document_syntax/heredoc_backslash_before_ordinary_character_retained
# In an unquoted here-document, a backslash before an ordinary character like 'n' or 't' is retained.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
doc=$(cat <<EOF
\n and \t remain literal
EOF
)
expected='\n and \t remain literal'
[ "$doc" = "$expected" ] || fail "backslash before ordinary character failed: got [$doc]"
echo PASS
exit 0
