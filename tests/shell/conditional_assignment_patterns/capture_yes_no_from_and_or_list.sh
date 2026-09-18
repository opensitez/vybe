#!/usr/bin/env bash
# vybe-test: bash/conditional_assignment_patterns/capture_yes_no_from_and_or_list
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=$( [ 1 -gt 0 ] && echo yes || echo no )
b=$( [ 1 -lt 0 ] && echo yes || echo no )
[ "$a" = yes ] && [ "$b" = no ] || fail "a=$a b=$b"
echo PASS
exit 0
