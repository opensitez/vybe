#!/usr/bin/env bash
# vybe-test: bash/conditional_assignment_patterns/assignment_status_drives_the_branch
# An assignment from a command substitution has that command's status, so it
# can be the condition itself.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if out=$(echo value); then r=ok; else r=bad; fi
[ "$r" = ok ] && [ "$out" = value ] || fail "success: r=$r out=[$out]"
if out=$(false); then r=ok; else r=bad; fi
[ "$r" = bad ] && [ -z "$out" ] || fail "failure: r=$r out=[$out]"
echo PASS
exit 0
