#!/usr/bin/env bash
# vybe-test: bash/compound_boolean_conditions/if_with_and_or_command_lists
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if true && [ 1 -eq 1 ]; then r=a; else r=b; fi
[ "$r" = a ] || fail "&& both true: got $r"
if true && false; then r=a; else r=b; fi
[ "$r" = b ] || fail "&& one false: got $r"
if false || [ 1 -eq 1 ]; then r=a; else r=b; fi
[ "$r" = a ] || fail "|| second true: got $r"
echo PASS
exit 0
