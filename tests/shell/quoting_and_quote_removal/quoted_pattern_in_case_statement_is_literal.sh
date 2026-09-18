#!/usr/bin/env bash
# vybe-test: bash/quoting_and_quote_removal/quoted_pattern_in_case_statement_is_literal
# Quoted characters in a case pattern are matched literally rather than as pattern wildcards.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
match1=""
case "abc" in
    "a*") match1="matched" ;;
    *) match1="missed" ;;
esac
[ "$match1" = "missed" ] || fail "quoted star in case should not match 'abc'"

match2=""
case "a*" in
    "a*") match2="matched" ;;
    *) match2="missed" ;;
esac
[ "$match2" = "matched" ] || fail "quoted star in case should match literal 'a*'"
echo PASS
exit 0
