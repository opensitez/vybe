#!/usr/bin/env bash
# vybe-test: bash/quoting_and_quote_removal/quoted_pattern_in_double_brackets_matches_literally
# Within [[ ... ]], quotes on the right-hand side of == force literal string matching instead of globbing.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
sample="hello*"
[[ "$sample" == "hello*" ]] || fail "literal quote pattern should match 'hello*'"
[[ "hello world" == "hello*" ]] && fail "quoted pattern should NOT match 'hello world'"
echo PASS
exit 0
