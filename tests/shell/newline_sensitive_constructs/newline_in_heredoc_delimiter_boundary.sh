#!/usr/bin/env bash
# vybe-test: bash/newline_sensitive_constructs/newline_in_heredoc_delimiter_boundary
# The delimiter token of a here-document must occupy an entire line terminated by a newline.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
doc=$(cat <<LIMIT
data line 1
data line 2
LIMIT
)
expected="data line 1"$'\n'"data line 2"
[ "$doc" = "$expected" ] || fail "heredoc newline delimiter boundary failed: got [$doc]"
echo PASS
exit 0
