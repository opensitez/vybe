#!/usr/bin/env bash
# vybe-test: bash/parser_contexts_and_ambiguity/command_substitution_containing_pipeline_and_subshell
# A complex construct combining subshells, pipelines, and command substitution parses unambiguously.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$( (printf 'input\n' | cat) | (cat) )
[ "$out" = "input" ] || fail "pipeline in nested subshells: want 'input', got [$out]"
echo PASS
exit 0
