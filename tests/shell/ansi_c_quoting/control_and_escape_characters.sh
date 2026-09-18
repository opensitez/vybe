#!/usr/bin/env bash
# vybe-test: bash/ansi_c_quoting/control_and_escape_characters
# \e is ESC (0x1b), \cX is control-X, \a is BEL, \r carriage return.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
esc=$'\e'; hex1b=$'\x1b'
[ "$esc" = "$hex1b" ] || fail "\\e is not 0x1b"
ca=$'\cA'; hex01=$'\x01'
[ "$ca" = "$hex01" ] || fail "\\cA is not 0x01"
cbr=$'\c['
[ "$cbr" = "$esc" ] || fail "\\c[ is not ESC"
bel=$'\a'; hex07=$'\x07'
[ "$bel" = "$hex07" ] || fail "\\a is not 0x07"
cr=$'\r'; hex0d=$'\x0d'
[ "$cr" = "$hex0d" ] || fail "\\r is not 0x0d"
echo PASS
exit 0
