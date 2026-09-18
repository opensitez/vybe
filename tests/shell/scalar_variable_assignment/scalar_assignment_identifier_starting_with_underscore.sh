#!/usr/bin/env bash
# vybe-test: bash/scalar_variable_assignment/scalar_assignment_identifier_starting_with_underscore
# Variable identifiers may begin with an underscore character '_'.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
_hidden_val="secret"
__deep_val="deeper"
[ "$_hidden_val" = "secret" ] || fail "_hidden_val assignment failed"
[ "$__deep_val" = "deeper" ] || fail "__deep_val assignment failed"
echo PASS
exit 0
