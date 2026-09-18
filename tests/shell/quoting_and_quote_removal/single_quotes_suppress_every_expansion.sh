#!/usr/bin/env bash
# vybe-test: bash/quoting_and_quote_removal/single_quotes_suppress_every_expansion
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=expanded
out='$x $(echo no) `echo no` \n * ~'
[ "$out" = '$x $(echo no) `echo no` \n * ~' ] || fail "got [$out]"
echo PASS
exit 0
