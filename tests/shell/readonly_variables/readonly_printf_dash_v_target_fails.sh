#!/usr/bin/env bash
# vybe-test: bash/readonly_variables/readonly_printf_dash_v_target_fails
# Attempting to assign into a readonly variable via 'printf -v' fails with non-zero exit status.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly BUF="const_buf"
( printf -v BUF "%s" "new_text" ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "printf -v targeting readonly variable should fail with non-zero exit code"
[ "$BUF" = "const_buf" ] || fail "readonly buffer was altered: got [$BUF]"
echo PASS
exit 0
