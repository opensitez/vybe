#!/usr/bin/env bash
# vybe-test: bash/name_resolution_and_identifiers/function_name_with_hyphens
# Bash allows function identifiers to contain hyphens as part of their name.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
kebab-case-helper() {
    printf 'invoked_kebab\n'
}
res=$(kebab-case-helper)
[ "$res" = "invoked_kebab" ] || fail "kebab function call: got [$res]"
echo PASS
exit 0
