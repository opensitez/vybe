#!/usr/bin/env bash
# vybe-test: bash/builtin_echo_option_parsing/repeat
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if (( 12 % 2 == 0 )); then
  out=$(builtin echo -n "ok12")
else
  out=$(builtin echo -- "--ok12")
fi
if (( 12 % 2 == 0 )); then
  [ "$out" = "ok12" ] || fail "echo -n option mismatch"
else
  [ "$out" = "--ok12" ] || fail "echo -- parsing mismatch"
fi
echo PASS
exit 0
