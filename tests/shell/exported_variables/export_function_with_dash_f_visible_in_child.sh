#!/usr/bin/env bash
# vybe-test: bash/exported_variables/export_function_with_dash_f_visible_in_child
# Exporting a function via 'export -f' makes it available to child Bash shell instances.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
exported_greeting() {
    printf 'hello_from_exported_fn\n'
}
export -f exported_greeting
child_out=$( "$BASH" -c 'exported_greeting' )
[ "$child_out" = "hello_from_exported_fn" ] || fail "exported function invocation in child failed: got [$child_out]"
echo PASS
exit 0
