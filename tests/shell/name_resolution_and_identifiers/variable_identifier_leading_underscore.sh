#!/usr/bin/env bash
# vybe-test: bash/name_resolution_and_identifiers/variable_identifier_leading_underscore
# Variable identifiers starting with an underscore are valid and distinct.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
_my_var="under"
__another="double_under"
[ "$_my_var" = "under" ] || fail "_my_var: want 'under', got [$_my_var]"
[ "$__another" = "double_under" ] || fail "__another: want 'double_under', got [$__another]"
echo PASS
exit 0
