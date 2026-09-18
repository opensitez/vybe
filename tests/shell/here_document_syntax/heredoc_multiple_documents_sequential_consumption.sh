#!/usr/bin/env bash
# vybe-test: bash/here_document_syntax/heredoc_multiple_documents_sequential_consumption
# Multiple here-documents can be attached to sequential commands within a compound group.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(
    { cat <<EOF1; cat <<EOF2; }
doc1_line
EOF1
doc2_line
EOF2
)
expected="doc1_line"$'\n'"doc2_line"
[ "$out" = "$expected" ] || fail "multiple heredocs sequential consumption failed: got [$out]"
echo PASS
exit 0
