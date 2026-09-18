#!/usr/bin/env bash
# vybe-test: bash/builtin_echo_option_parsing/redefine
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if (( 9 % 2 == 0 )); then
  out=$(builtin echo -n "ok9")
else
  out=$(builtin echo -- "--ok9")
fi
if (( 9 % 2 == 0 )); then
  [ "$out" = "ok9" ] || fail "echo -n option mismatch"
else
  [ "$out" = "--ok9" ] || fail "echo -- parsing mismatch"
fi
echo PASS
exit 0
