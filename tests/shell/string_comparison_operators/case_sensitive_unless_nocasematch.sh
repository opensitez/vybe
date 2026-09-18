#!/usr/bin/env bash
# vybe-test: bash/string_comparison_operators/case_sensitive_unless_nocasematch
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ ABC == abc ]] && fail "default must be case sensitive"
shopt -s nocasematch
[[ ABC == abc ]] || fail "nocasematch must make [[ case insensitive"
[ ABC = abc ] && fail "[ is not affected by nocasematch"
echo PASS
exit 0
