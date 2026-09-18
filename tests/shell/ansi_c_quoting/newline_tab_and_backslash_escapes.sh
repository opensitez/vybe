#!/usr/bin/env bash
# vybe-test: bash/ansi_c_quoting/newline_tab_and_backslash_escapes
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=$'a\tb\nc'
[ "${#x}" -eq 5 ] || fail "length want 5 got ${#x}"
[ "${x:1:1}" = "	" ] || fail "second char must be a tab"
bs=$'\\'
[ "$bs" = '\' ] && [ "${#bs}" -eq 1 ] || fail "double backslash must be one backslash"
echo PASS
exit 0
