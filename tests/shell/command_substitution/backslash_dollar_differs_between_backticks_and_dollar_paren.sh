#!/usr/bin/env bash
# vybe-test: bash/command_substitution/backslash_dollar_differs_between_backticks_and_dollar_paren
# Inside backticks a backslash before $ ` or \ is removed before the inner
# command is parsed, so \$x is expanded; inside $( ) it stays an escape.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=val
bt=`echo \$x`
[ "$bt" = val ] || fail "backticks: want [val] got [$bt]"
dp=$(echo \$x)
[ "$dp" = '$x' ] || fail "\$( ): want [\$x] got [$dp]"
echo PASS
exit 0
