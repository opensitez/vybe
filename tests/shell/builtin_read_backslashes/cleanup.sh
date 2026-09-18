#!/usr/bin/env bash
# vybe-test: bash/builtin_read_backslashes/cleanup
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
src='a\ b\ c'
if (( 15 % 2 == 0 )); then
  read out <<< "$src"
  [ "$out" = "a b c" ] || fail "unquoted read should unescape spaces"
else
  read -r out <<< "$src"
  [ "$out" = 'a\ b\ c' ] || fail "read -r should preserve escapes"
fi
echo PASS
exit 0
