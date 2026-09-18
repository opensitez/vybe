#!/usr/bin/env bash
# vybe-test: bash/parentheses_and_nested_commands/parentheses_case_pattern_optional_open_paren
# In case statements, an optional opening parenthesis '(' can precede each pattern before ')'.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
match1=""
case "opt" in
    (opt) match1="with_open" ;;
    *) match1="miss" ;;
esac
[ "$match1" = "with_open" ] || fail "pattern with open paren failed"

match2=""
case "alt" in
    (alt|other) match2="matched_alt" ;;
    *) match2="miss" ;;
esac
[ "$match2" = "matched_alt" ] || fail "alternation with open paren failed"
echo PASS
exit 0
