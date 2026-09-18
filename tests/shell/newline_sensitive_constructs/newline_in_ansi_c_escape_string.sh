#!/usr/bin/env bash
# vybe-test: bash/newline_sensitive_constructs/newline_in_ansi_c_escape_string
# The $'\n' construct produces an exact ASCII newline byte (hex 0x0A).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
nl=$'\n'
[ "${#nl}" -eq 1 ] || fail "newline length: want 1, got ${#nl}"
printf -v hex '%02x' "'$nl"
[ "$hex" = "0a" ] || fail "newline byte mismatch: want hex '0a', got [$hex]"
echo PASS
exit 0
