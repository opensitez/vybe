#!/usr/bin/env bash
# vybe-test: bash/backslash_escaping/backslash_escapes_tilde_expansion
# A backslash before '~' suppresses tilde expansion to the home directory.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tilde_escaped=\~
[ "$tilde_escaped" = "~" ] || fail "escaped tilde: want '~', got [$tilde_escaped]"
echo PASS
exit 0
