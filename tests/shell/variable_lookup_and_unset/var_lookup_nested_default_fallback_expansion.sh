#!/usr/bin/env bash
# vybe-test: bash/variable_lookup_and_unset/var_lookup_nested_default_fallback_expansion
# Default parameter expansions can nest ${v1:-${v2:-fallback}} to form fallback chains.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset level1 level2
res1="${level1:-${level2:-ultimate_fallback}}"
[ "$res1" = "ultimate_fallback" ] || fail "deep fallback failed: got [$res1]"

level2="second_level"
res2="${level1:-${level2:-ultimate_fallback}}"
[ "$res2" = "second_level" ] || fail "intermediate fallback failed: got [$res2]"

level1="first_level"
res3="${level1:-${level2:-ultimate_fallback}}"
[ "$res3" = "first_level" ] || fail "primary value failed: got [$res3]"
echo PASS
exit 0
