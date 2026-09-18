#!/usr/bin/env bash
# vybe-test: bash/command_substitution/null_bytes_are_dropped_with_a_warning
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$( { x=$(printf 'a\0b'); echo "value=$x"; } 2>&1 )
[[ $msg == *"ignored null byte"* ]] || fail "want warning got [$msg]"
[[ $msg == *"value=ab" ]] || fail "want value=ab got [$msg]"
echo PASS
exit 0
