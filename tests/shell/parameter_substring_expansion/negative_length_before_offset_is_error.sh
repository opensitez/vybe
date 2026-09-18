#!/usr/bin/env bash
# vybe-test: bash/parameter_substring_expansion/negative_length_before_offset_is_error
# When the end computed from a negative length lies before the offset the
# expansion fails with "substring expression < 0" and aborts the subshell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=abc
msg=$( eval 'echo "${x:1:-5}"; echo unreachable' 2>&1 ); st=$?
[ "$st" -eq 1 ] || fail "want status 1 got $st"
[[ $msg == *"substring expression < 0"* ]] || fail "got [$msg]"
[[ $msg != *unreachable* ]] || fail "subshell must abort"
echo PASS
exit 0
