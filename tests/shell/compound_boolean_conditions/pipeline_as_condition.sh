#!/usr/bin/env bash
# vybe-test: bash/compound_boolean_conditions/pipeline_as_condition
# The pipeline's status (its last command) is the condition.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if echo x | { read -r v; [ "$v" = x ]; }; then r=yes; else r=no; fi
[ "$r" = yes ] || fail "matching read: got $r"
if false | true; then r=yes; else r=no; fi
[ "$r" = yes ] || fail "last command true: got $r"
if true | false; then r=yes; else r=no; fi
[ "$r" = no ] || fail "last command false: got $r"
echo PASS
exit 0
