#!/usr/bin/env bash
# vybe-test: bash/here_string_syntax/here_string_expands_tilde_path
# Tilde expansion occurs on unquoted here-string arguments starting with '~'.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
read -r result <<< ~
[ "$result" = "$HOME" ] || fail "tilde expansion in here-string: want [$HOME], got [$result]"
echo PASS
exit 0
