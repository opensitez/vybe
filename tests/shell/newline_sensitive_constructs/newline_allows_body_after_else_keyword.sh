#!/usr/bin/env bash
# vybe-test: bash/newline_sensitive_constructs/newline_allows_body_after_else_keyword
# In an if statement, a newline after 'else' allows body commands on subsequent lines without a semicolon.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
branch="none"
if false; then
    branch="if"
else
    branch="else_executed"
fi
[ "$branch" = "else_executed" ] || fail "body after else newline failed: got [$branch]"
echo PASS
exit 0
