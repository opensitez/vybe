#!/usr/bin/env bash
# vybe-test: bash/process_substitution/loop_reading_from_substitution_keeps_variables
# Unlike a pipeline, a loop fed by < <(list) runs in the current shell, so
# variables set inside it survive.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
n=0
while read -r _; do n=$((n+1)); done < <(printf 'a\nb\nc\n')
[ "$n" -eq 3 ] || fail "want 3 got $n"
m=0
printf 'a\nb\n' | while read -r _; do m=$((m+1)); done
[ "$m" -eq 0 ] || fail "pipeline loop runs in a subshell, want 0 got $m"
echo PASS
exit 0
