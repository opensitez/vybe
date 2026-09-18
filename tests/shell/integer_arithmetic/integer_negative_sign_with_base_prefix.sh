#!/usr/bin/env bash
# vybe-test: bash/integer_arithmetic/integer_negative_sign_with_base_prefix
# The base prefix follows the sign only in arithmetic parsing.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((-2#111)) -eq -7 ] || fail "-2#111 got $((-2#111))"
[ $((-16#ff + 16#ff)) -eq 0 ] || fail "-16#ff+16#ff got $((-16#ff + 16#ff))"
echo PASS
exit 0
