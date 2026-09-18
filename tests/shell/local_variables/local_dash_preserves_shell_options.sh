#!/usr/bin/env bash
# vybe-test: bash/local_variables/local_dash_preserves_shell_options
# The 'local -' statement saves shell options locally, restoring original options when the function returns.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set +e
option_altering_fn() {
    local -
    set -e
    case "$-" in
        *e*) : ;;
        *) exit 1 ;;
    esac
}
option_altering_fn
case "$-" in
    *e*) fail "set -e leaked beyond function scope despite 'local -'" ;;
esac
echo PASS
exit 0
