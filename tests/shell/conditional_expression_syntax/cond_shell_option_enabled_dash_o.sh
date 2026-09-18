#!/usr/bin/env bash
# vybe-test: bash/conditional_expression_syntax/cond_shell_option_enabled_dash_o
# The -o optname operator tests whether a specific shell option is enabled.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -e
[[ -o errexit ]] || fail "errexit should test true with -o after set -e"
set +e
[[ -o errexit ]] && fail "errexit should test false with -o after set +e"
echo PASS
exit 0
