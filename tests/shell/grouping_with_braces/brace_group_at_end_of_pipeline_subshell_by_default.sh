#!/usr/bin/env bash
# vybe-test: bash/grouping_with_braces/brace_group_at_end_of_pipeline_subshell_by_default
# By default, each component of a pipeline (including a brace group) executes in a subshell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -u lastpipe 2>/dev/null
captured="before"
printf 'new_value\n' | { read -r captured; }
[ "$captured" = "before" ] || fail "pipeline element should run in subshell by default: got [$captured]"
echo PASS
exit 0
