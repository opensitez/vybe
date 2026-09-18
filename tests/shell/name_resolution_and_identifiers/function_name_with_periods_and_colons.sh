#!/usr/bin/env bash
# vybe-test: bash/name_resolution_and_identifiers/function_name_with_periods_and_colons
# In Bash, function names can contain characters like periods and colons, unlike variable names.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
module.submodule:action() {
    printf 'namespaced_action\n'
}
res=$(module.submodule:action)
[ "$res" = "namespaced_action" ] || fail "namespaced function call: got [$res]"
echo PASS
exit 0
