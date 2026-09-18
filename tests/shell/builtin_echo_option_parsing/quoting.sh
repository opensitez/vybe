#!/usr/bin/env bash
# vybe-test: bash/builtin_echo_option_parsing/quoting
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if (( 6 % 2 == 0 )); then
  out=$(builtin echo -n "ok6")
else
  out=$(builtin echo -- "--ok6")
fi
if (( 6 % 2 == 0 )); then
  [ "$out" = "ok6" ] || fail "echo -n option mismatch"
else
  [ "$out" = "--ok6" ] || fail "echo -- parsing mismatch"
fi
echo PASS
exit 0
