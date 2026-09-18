#!/usr/bin/env bash
# vybe-test: bash/parameter_length_expansion/param_len_utf8_multibyte_characters
# The ${#var} expansion counts UTF-8 characters (codepoints), not raw bytes.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
jp_str="日本語"
[ "${#jp_str}" -eq 3 ] || fail "Japanese character length: want 3, got ${#jp_str}"

accent_str="café"
[ "${#accent_str}" -eq 4 ] || fail "Accented character length: want 4, got ${#accent_str}"
echo PASS
exit 0
