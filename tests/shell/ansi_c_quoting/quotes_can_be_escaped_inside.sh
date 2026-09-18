#!/usr/bin/env bash
# vybe-test: bash/ansi_c_quoting/quotes_can_be_escaped_inside
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=$'it\'s'
[ "$a" = "it's" ] || fail "escaped single quote: got [$a]"
b=$'say \"hi\"'
[ "$b" = 'say "hi"' ] || fail "escaped double quote: got [$b]"
c=$'plain "quotes" ok'
[ "$c" = 'plain "quotes" ok' ] || fail "unescaped double quotes are literal: got [$c]"
echo PASS
exit 0
