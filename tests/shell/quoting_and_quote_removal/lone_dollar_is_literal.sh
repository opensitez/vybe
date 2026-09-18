#!/usr/bin/env bash
# vybe-test: bash/quoting_and_quote_removal/lone_dollar_is_literal
# A $ not followed by a name, digit, brace or paren expands to itself.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(echo "$" a$ '$')
[ "$out" = '$ a$ $' ] || fail "want [\$ a\$ \$] got [$out]"
echo PASS
exit 0
