#!/usr/bin/env bash
# vybe-test: bash/readonly_variables/readonly_declare_dash_r_syntax
# 'declare -r' is equivalent to 'readonly' and produces an immutable variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -r SECURE_KEY="secret_key_123"
[ "$SECURE_KEY" = "secret_key_123" ] || fail "declare -r failed: got [$SECURE_KEY]"
( SECURE_KEY="new" ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "declare -r variable should reject reassignment"
echo PASS
exit 0
