#!/usr/bin/env bash
# vybe-test: bash/variable_attributes_and_introspection/intro_prefix_expansion_star_introspects_matching_names
# The ${!prefix*} syntax introspects variable names by prefix returning a space-separated string.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
INTRO_STAR_1=10
INTRO_STAR_2=20
matches="${!INTRO_STAR_*}"
case "$matches" in
    *"INTRO_STAR_1"*"INTRO_STAR_2"*|*"INTRO_STAR_2"*"INTRO_STAR_1"*) : ;;
    *) fail "prefix star expansion failed to introspect variables: got [$matches]" ;;
esac
echo PASS
exit 0
