#!/usr/bin/env bash
# vybe-test: bash/parameter_error_substitutions/param_err_assignment_combined_syntax
# Error substitution can be combined with variable assignment; the destination receives the value on success.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
source_var="established_value"
dest_var="${source_var:?cannot be empty}"
[ "$dest_var" = "established_value" ] || fail "assignment from :? expression failed: got [$dest_var]"
echo PASS
exit 0
