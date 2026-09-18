#!/usr/bin/env bash
# vybe-test: bash/local_variables/local_outside_function_returns_error
# Invoking the 'local' builtin outside of a function body produces a non-zero exit error.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval 'local illegal_var="value"' 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "local outside function should return non-zero exit code"
echo PASS
exit 0
