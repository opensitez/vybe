#!/usr/bin/env bash
# vybe-test: bash/here_document_syntax/heredoc_unquoted_expands_arithmetic_evaluation
# In an unquoted here-document, $((...)) arithmetic evaluations are calculated and inserted.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=15
y=25
doc=$(cat <<EOF
Total: $(( x + y ))
EOF
)
[ "$doc" = "Total: 40" ] || fail "arithmetic in heredoc failed: got [$doc]"
echo PASS
exit 0
