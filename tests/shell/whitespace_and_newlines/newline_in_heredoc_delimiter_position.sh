#!/usr/bin/env bash
# vybe-test: bash/whitespace_and_newlines/newline_in_heredoc_delimiter_position
# The delimiter of a here-document must be immediately followed by a newline.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
content=$(cat <<'EOF'
line one
line two
EOF
)
expected="line one"$'\n'"line two"
[ "$content" = "$expected" ] || fail "heredoc content mismatch: got [$content]"
echo PASS
exit 0
