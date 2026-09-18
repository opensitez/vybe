#!/usr/bin/env bash
# vybe-test: bash/case_pattern_syntax/case_pattern_clause_terminator_double_semicolon
# The ';;' operator terminates execution of the case statement after the matching clause executes.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
trail=""
case "first" in
    first)
        trail+="1"
        ;;
    *)
        trail+="2"
        ;;
esac
[ "$trail" = "1" ] || fail ";; failed to terminate case execution: got [$trail]"
echo PASS
exit 0
