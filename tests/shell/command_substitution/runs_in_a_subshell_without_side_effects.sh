#!/usr/bin/env bash
# vybe-test: bash/command_substitution/runs_in_a_subshell_without_side_effects
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
v=1; here=$PWD
_=$(v=2; cd /; g() { :; }; set -- z; shift)
[ "$v" = 1 ] || fail "variable leaked: v=$v"
[ "$PWD" = "$here" ] || fail "cd leaked"
[ "$(type -t g)" = "" ] || fail "function definition leaked"
echo PASS
exit 0
