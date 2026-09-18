#!/usr/bin/env bash
# vybe-test: bash/variable_lookup_and_unset/var_lookup_unset_variable_expands_to_empty
# Under default shell options, referencing an unset variable expands to an empty string.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset missing_var
val="prefix_${missing_var}_suffix"
[ "$val" = "prefix__suffix" ] || fail "unset expansion: want 'prefix__suffix', got [$val]"
echo PASS
exit 0
