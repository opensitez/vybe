#!/usr/bin/env bash
# vybe-test: bash/variable_attributes_and_introspection/intro_compgen_dash_v_discovers_variable_names_by_prefix
# The 'compgen -v prefix' builtin dynamically introspects and lists all variables matching prefix.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
INTRO_DISCOVERY_ALPHA=1
INTRO_DISCOVERY_BETA=2
found=$(compgen -v INTRO_DISCOVERY_)
case "$found" in
    *"INTRO_DISCOVERY_ALPHA"*"INTRO_DISCOVERY_BETA"*|*"INTRO_DISCOVERY_BETA"*"INTRO_DISCOVERY_ALPHA"*) : ;;
    *) fail "compgen -v failed to discover matching variables: got [$found]" ;;
esac
echo PASS
exit 0
