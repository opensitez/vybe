#!/usr/bin/env bash
# vybe-test: bash/ansi_c_quoting/unicode_escapes_produce_utf8_characters
# \uHHHH and \UHHHHHHHH are encoded in the current locale's charset; under
# UTF-8 the result is one character, under C it is counted as bytes.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export LC_ALL=en_US.UTF-8
e=$'é'
[ "$e" = 'é' ] || fail "\\u00e9 want é got [$e]"
[ "${#e}" -eq 1 ] || fail "one character under UTF-8, got ${#e}"
smile=$'\U0001F600'
[ "${#smile}" -eq 1 ] || fail "\\U escape is one character, got ${#smile}"
export LC_ALL=C
[ "${#e}" -eq 2 ] || fail "é is 2 bytes under C, got ${#e}"
[ "${#smile}" -eq 4 ] || fail "U+1F600 is 4 bytes, got ${#smile}"
echo PASS
exit 0
