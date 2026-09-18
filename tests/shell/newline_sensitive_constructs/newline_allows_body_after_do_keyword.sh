#!/usr/bin/env bash
# vybe-test: bash/newline_sensitive_constructs/newline_allows_body_after_do_keyword
# In while and for loops, a newline after 'do' allows loop body commands without a semicolon.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
acc=""
for item in x y z
do
    acc+="$item"
done
[ "$acc" = "xyz" ] || fail "body after do newline: want 'xyz', got [$acc]"
echo PASS
exit 0
