#!/usr/bin/env bash
# vybe-test: bash/variable_lookup_and_unset/var_lookup_quoted_preserves_internal_spaces
# Referencing "$var" within double quotes preserves internal spaces and treats the value as one word.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
spaced="one   two   three"
count_args() { printf '%s\n' "$#"; }
c=$(count_args "$spaced")
[ "$c" -eq 1 ] || fail "quoted expansion argument count: want 1, got $c"
echo PASS
exit 0
