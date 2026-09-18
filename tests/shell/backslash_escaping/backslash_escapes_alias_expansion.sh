#!/usr/bin/env bash
# vybe-test: bash/backslash_escaping/backslash_escapes_alias_expansion
# Prepending a backslash to a command name suppresses alias lookup for that command.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s expand_aliases
alias custom_echo="printf 'aliased\n'"
out1=$(custom_echo)
out2=$(\custom_echo 2>/dev/null || printf 'not_aliased\n')
[ "$out1" = "aliased" ] || fail "out1 should be aliased, got [$out1]"
[ "$out2" = "not_aliased" ] || fail "out2 should bypass alias, got [$out2]"
echo PASS
exit 0
