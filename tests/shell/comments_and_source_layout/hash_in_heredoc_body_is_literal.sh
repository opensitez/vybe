#!/usr/bin/env bash
# vybe-test: bash/comments_and_source_layout/hash_in_heredoc_body_is_literal
# A line beginning with '#' inside a here-document is literal input, NOT a shell comment.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
doc=$(cat <<'EOF'
# not a comment
data line
EOF
)
expected="# not a comment"$'\n'"data line"
[ "$doc" = "$expected" ] || fail "heredoc hash mismatch: got [$doc]"
echo PASS
exit 0
