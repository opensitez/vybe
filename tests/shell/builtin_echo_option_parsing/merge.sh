#!/usr/bin/env bash
# vybe-test: bash/builtin_echo_option_parsing/merge
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if (( 14 % 2 == 0 )); then
  out=$(builtin echo -n "ok14")
else
  out=$(builtin echo -- "--ok14")
fi
if (( 14 % 2 == 0 )); then
  [ "$out" = "ok14" ] || fail "echo -n option mismatch"
else
  [ "$out" = "--ok14" ] || fail "echo -- parsing mismatch"
fi
echo PASS
exit 0
