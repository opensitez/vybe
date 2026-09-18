#!/usr/bin/env bash
# vybe-test: bash/parameter_pattern_replacement/empty_anchored_pattern_prepends_or_appends
# An empty pattern anchored with # or % matches the empty string at that edge.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=abc
[ "${x/#/pre-}" = pre-abc ] || fail "prepend: got [${x/#/pre-}]"
[ "${x/%/-suf}" = abc-suf ] || fail "append: got [${x/%/-suf}]"
echo PASS
exit 0
