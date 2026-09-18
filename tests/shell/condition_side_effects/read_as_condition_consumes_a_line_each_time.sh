#!/usr/bin/env bash
# vybe-test: bash/condition_side_effects/read_as_condition_consumes_a_line_each_time
# The failing final read leaves the variable empty.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
n=0
while read -r line; do n=$((n+1)); last=$line; done <<< $'a\nb'
[ "$n" -eq 2 ] || fail "iterations want 2 got $n"
[ "$last" = b ] || fail "last successful line got [$last]"
[ -z "$line" ] || fail "variable is emptied by the failing read, got [$line]"
echo PASS
exit 0
