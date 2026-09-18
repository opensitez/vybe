#!/usr/bin/env bash
# vybe-test: bash/parser_contexts_and_ambiguity/nested_command_substitutions_ambiguity_resolution
# Deeply nested command substitutions $(... $(...)...) resolve their boundaries cleanly.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
res=$(echo $(echo $(echo "deeply_nested")))
[ "$res" = "deeply_nested" ] || fail "deeply nested cmdsub: want 'deeply_nested', got [$res]"
echo PASS
exit 0
