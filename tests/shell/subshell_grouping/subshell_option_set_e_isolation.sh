#!/usr/bin/env bash
# vybe-test: bash/subshell_grouping/subshell_option_set_e_isolation
# Modifying shell options like 'set -e' or 'set -u' inside a subshell does not alter the parent shell options.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set +e
(
    set -e
    # Inside subshell, $- contains e
    case "$-" in
        *e*) exit 0 ;;
        *) exit 1 ;;
    esac
)
st=$?
[ "$st" -eq 0 ] || fail "subshell set -e failed"
case "$-" in
    *e*) fail "parent shell option -e was unexpectedly enabled" ;;
esac
echo PASS
exit 0
