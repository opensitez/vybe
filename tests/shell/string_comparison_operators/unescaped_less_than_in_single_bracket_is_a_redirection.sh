#!/usr/bin/env bash
# vybe-test: bash/string_comparison_operators/unescaped_less_than_in_single_bracket_is_a_redirection
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$( { [ a < nonexistent_file_zz ]; } 2>&1 ); st=$?
[ "$st" -eq 1 ] || fail "want status 1 got $st"
[[ $msg == *"No such file"* ]] || fail "got [$msg]"
echo PASS
exit 0
