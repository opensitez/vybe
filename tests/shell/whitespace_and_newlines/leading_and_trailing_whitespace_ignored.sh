#!/usr/bin/env bash
# vybe-test: bash/whitespace_and_newlines/leading_and_trailing_whitespace_ignored
# Leading and trailing spaces and tabs around command invocations are discarded.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
   val=42   
[ "$val" -eq 42 ] || fail "val with surrounding whitespace: want 42, got $val"
		val=84		
[ "$val" -eq 84 ] || fail "val with surrounding tabs: want 84, got $val"
echo PASS
exit 0
