#!/usr/bin/env bash
# vybe-test: bash/builtin_echo_option_parsing/override
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if (( 4 % 2 == 0 )); then
  out=$(builtin echo -n "ok4")
else
  out=$(builtin echo -- "--ok4")
fi
if (( 4 % 2 == 0 )); then
  [ "$out" = "ok4" ] || fail "echo -n option mismatch"
else
  [ "$out" = "--ok4" ] || fail "echo -- parsing mismatch"
fi
echo PASS
exit 0
