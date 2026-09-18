#!/usr/bin/env bash
# vybe-test: bash/ansi_c_quoting/unknown_escape_is_kept_literally
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=$'\z\q'
[ "$x" = '\z\q' ] || fail "got [$x]"
[ "${#x}" -eq 4 ] || fail "length want 4 got ${#x}"
echo PASS
exit 0
