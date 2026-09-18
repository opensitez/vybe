#!/usr/bin/env bash
# vybe-test: bash/variable_lookup_and_unset/var_lookup_unquoted_performs_glob_expansion
# Referencing unquoted $var where the value contains glob characters expands to matching filesystem paths.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
pat="tests/bash/*"
set -- $pat
[ "$#" -gt 5 ] || fail "unquoted glob expansion failed to match directory entries: got count $#"
echo PASS
exit 0
