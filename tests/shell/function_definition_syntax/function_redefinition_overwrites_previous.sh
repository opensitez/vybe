#!/usr/bin/env bash
# vybe-test: bash/function_definition_syntax/function_redefinition_overwrites_previous
# Defining a function with the same name overwrites the prior definition completely.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
dynamic_fn() { printf 'version_1\n'; }
[ "$(dynamic_fn)" = "version_1" ] || fail "initial function definition failed"

dynamic_fn() { printf 'version_2\n'; }
[ "$(dynamic_fn)" = "version_2" ] || fail "redefined function failed"
echo PASS
exit 0
