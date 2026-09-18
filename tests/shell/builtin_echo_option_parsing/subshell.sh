#!/usr/bin/env bash
# vybe-test: bash/builtin_echo_option_parsing/subshell
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if (( 3 % 2 == 0 )); then
  out=$(builtin echo -n "ok3")
else
  out=$(builtin echo -- "--ok3")
fi
if (( 3 % 2 == 0 )); then
  [ "$out" = "ok3" ] || fail "echo -n option mismatch"
else
  [ "$out" = "--ok3" ] || fail "echo -- parsing mismatch"
fi
echo PASS
exit 0
