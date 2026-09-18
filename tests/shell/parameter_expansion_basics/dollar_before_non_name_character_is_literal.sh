#!/usr/bin/env bash
# vybe-test: bash/parameter_expansion_basics/dollar_before_non_name_character_is_literal
# $ followed by a character that cannot start a parameter (%, space, end of
# word, a quote) is an ordinary dollar sign.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(echo "$%" "$ x" "a$" '$'"'")
[ "$out" = "\$% \$ x a\$ \$'" ] || fail "got [$out]"
echo PASS
exit 0
