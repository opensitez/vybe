#!/usr/bin/env bash
# vybe-test: bash/here_document_syntax/heredoc_strip_tabs_preserves_leading_spaces
# The <<- operator strips ONLY leading tabs, never leading spaces.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
doc=$(cat <<-EOF
    four_spaces
  two_spaces
EOF
)
expected="    four_spaces"$'\n'"  two_spaces"
[ "$doc" = "$expected" ] || fail "leading spaces altered by <<-"
echo PASS
exit 0
