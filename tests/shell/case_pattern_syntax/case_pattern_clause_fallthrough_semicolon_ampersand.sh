#!/usr/bin/env bash
# vybe-test: bash/case_pattern_syntax/case_pattern_clause_fallthrough_semicolon_ampersand
# The ';&' operator falls through directly to the command list of the subsequent clause without pattern check.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
seq=""
case "start" in
    start)
        seq+="step1,"
        ;&
    unmatched_pattern)
        seq+="step2,"
        ;&
    final)
        seq+="step3"
        ;;
esac
[ "$seq" = "step1,step2,step3" ] || fail ";& fallthrough chain failed: got [$seq]"
echo PASS
exit 0
