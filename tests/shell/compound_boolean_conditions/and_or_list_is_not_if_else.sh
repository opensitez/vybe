#!/usr/bin/env bash
# vybe-test: bash/compound_boolean_conditions/and_or_list_is_not_if_else
# a && b || c runs c whenever b fails, even though a succeeded.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
ran=
true && false || ran=c
[ "$ran" = c ] || fail "c must run when b fails"
ran=
if true; then false; else ran=c; fi
[ -z "$ran" ] || fail "real if/else does not run the else branch"
echo PASS
exit 0
