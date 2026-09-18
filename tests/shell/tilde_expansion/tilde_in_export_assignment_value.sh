#!/usr/bin/env bash
# vybe-test: bash/tilde_expansion/tilde_in_export_assignment_value
# The export builtin expands tildes in its assignment arguments.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ -n "$HOME" ] || fail "HOME is not set"
export EXPORTED_DIR=~/export_data
[ "$EXPORTED_DIR" = "$HOME/export_data" ] || fail "export tilde failed: want [$HOME/export_data], got [$EXPORTED_DIR]"
echo PASS
exit 0
