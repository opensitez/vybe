#!/usr/bin/env bash
# vybe-test: bash/tilde_expansion/tilde_in_declare_assignment_value
# The declare builtin expands tildes in its assignment arguments.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ -n "$HOME" ] || fail "HOME is not set"
declare declared_path=~/declared_dir
[ "$declared_path" = "$HOME/declared_dir" ] || fail "declare tilde failed: want [$HOME/declared_dir], got [$declared_path]"
echo PASS
exit 0
