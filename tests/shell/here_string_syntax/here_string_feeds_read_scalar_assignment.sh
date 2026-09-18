#!/usr/bin/env bash
# vybe-test: bash/here_string_syntax/here_string_feeds_read_scalar_assignment
# Passing a here-string directly to read assigns words into multiple scalar variables based on IFS.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
read -r w1 w2 w3 <<< "apple banana cherry"
[ "$w1" = "apple" ] || fail "w1: want 'apple', got [$w1]"
[ "$w2" = "banana" ] || fail "w2: want 'banana', got [$w2]"
[ "$w3" = "cherry" ] || fail "w3: want 'cherry', got [$w3]"
echo PASS
exit 0
