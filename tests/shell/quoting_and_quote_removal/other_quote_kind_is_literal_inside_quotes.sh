#!/usr/bin/env bash
# vybe-test: bash/quoting_and_quote_removal/other_quote_kind_is_literal_inside_quotes
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a="it's"
[ "$a" = "it's" ] || fail "single inside double: got [$a]"
b='say "hi"'
[ "$b" = 'say "hi"' ] || fail "double inside single: got [$b]"
[ "${#b}" -eq 8 ] || fail "length of say \"hi\" want 8 got ${#b}"
echo PASS
exit 0
