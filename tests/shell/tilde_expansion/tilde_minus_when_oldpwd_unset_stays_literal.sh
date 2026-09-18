#!/usr/bin/env bash
# vybe-test: bash/tilde_expansion/tilde_minus_when_oldpwd_unset_stays_literal
# When OLDPWD is unset, '~-' is not expanded and remains as the literal string '~-'.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset OLDPWD
res=~-
[ "$res" = "~-" ] || fail "unset OLDPWD should leave '~-' literal: got [$res]"
echo PASS
exit 0
