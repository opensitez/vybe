#!/usr/bin/env bash
# vybe-test: bash/builtin_command_dispatch/regex_capture
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=14
printf_fn(){ :; }
printf(){ return 7; }
out=$(command printf '%s' "dispatch_${IDX}")
[[ $out == "dispatch_${IDX}" ]] || fail "command builtin should bypass function named printf"
unset -f printf
if (( IDX % 2 == 0 )); then
  command printf '%s' "dispatch_reset_${IDX}" >/dev/null
fi
echo PASS
exit 0
