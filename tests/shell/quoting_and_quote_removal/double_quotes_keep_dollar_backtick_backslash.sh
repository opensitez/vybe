#!/usr/bin/env bash
# vybe-test: bash/quoting_and_quote_removal/double_quotes_keep_dollar_backtick_backslash
# Inside "…" parameter, command and arithmetic expansion still happen, but
# glob and tilde do not.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=v
out="$x $(echo c) `echo b` $((1+1)) * ~"
[ "$out" = 'v c b 2 * ~' ] || fail "got [$out]"
echo PASS
exit 0
