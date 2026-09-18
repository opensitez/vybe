#!/usr/bin/env bash
# vybe-test: bash/quoting_and_quote_removal/ansi_c_hex_and_octal_escapes
# ANSI-C quoting $'\xhh' and $'\ooo' resolves hex and octal byte escapes.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
s_hex=$'\x41\x42\x43'
s_oct=$'\101\102\103'
[ "$s_hex" = "ABC" ] || fail "hex escape: want 'ABC', got [$s_hex]"
[ "$s_oct" = "ABC" ] || fail "octal escape: want 'ABC', got [$s_oct]"
echo PASS
exit 0
