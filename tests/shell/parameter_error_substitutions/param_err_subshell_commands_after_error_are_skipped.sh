#!/usr/bin/env bash
# vybe-test: bash/parameter_error_substitutions/param_err_subshell_commands_after_error_are_skipped
# Following a fatal parameter error substitution, subsequent subshell commands are not executed.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset needle
side_effect_triggered=0
(
    val="${needle:?halt}"
    side_effect_triggered=1
) 2>/dev/null
[ "$side_effect_triggered" -eq 0 ] || fail "command executed after parameter error"
echo PASS
exit 0
