#!/usr/bin/env bash
# vybe-test: bash/newline_sensitive_constructs/newline_allows_body_after_then_keyword
# In an if statement, a newline after 'then' eliminates the need for a semicolon before the body.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
res="none"
if true
then
    res="then_executed"
fi
[ "$res" = "then_executed" ] || fail "body after then newline failed: got [$res]"
echo PASS
exit 0
