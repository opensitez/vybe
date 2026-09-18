#!/usr/bin/env bash
# vybe-test: bash/variable_scope_and_shadowing/scope_pipeline_components_isolate_modifications
# Variables modified inside pipeline stages run in subshells and isolate modifications from the caller.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
pipe_scope_var="unmodified"
printf 'line\n' | { pipe_scope_var="piped_mod"; }
[ "$pipe_scope_var" = "unmodified" ] || fail "pipeline element leaked modification to outer scope: got [$pipe_scope_var]"
echo PASS
exit 0
