#!/usr/bin/env bash
# vybe-test: bash/syntax_errors_and_recovery/eval_parses_whole_string_before_running
# eval parses its entire argument first: a syntax error anywhere means no
# command in the string executes, not even the ones before the bad token.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(eval 'echo first; echo a ) b; echo third' 2>/dev/null); st=$?
[ "$st" -ne 0 ] || fail "eval must fail"
[ -z "$out" ] || fail "nothing may run, got [$out]"
echo PASS
exit 0
