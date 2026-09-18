#!/usr/bin/env bash
# vybe-test: bash/quoting_and_quote_removal/quoted_command_name_still_runs
# Quote removal happens before command lookup; only reserved words are affected
# by quoting, builtins and functions are not.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
greet() { echo hi; }
out=$("ech""o" one)
[ "$out" = one ] || fail "builtin: want [one] got [$out]"
out=$('gre'et)
[ "$out" = hi ] || fail "function: want [hi] got [$out]"
echo PASS
exit 0
