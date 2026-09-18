#!/usr/bin/env bash
# vybe-test: bash/positional_parameters/pos_param_set_double_dash_clears_all
# Invoking 'set --' without any arguments clears all positional parameters, setting $# to 0.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "arg1" "arg2"
[ "$#" -eq 2 ] || fail "initial set failed"
set --
[ "$#" -eq 0 ] || fail "parameter count after 'set --': want 0, got $#"
[ -z "$1" ] || fail "\$1 should expand to empty after clearing: got [$1]"
echo PASS
exit 0
