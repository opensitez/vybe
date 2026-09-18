#!/usr/bin/env bash
# vybe-test: bash/case_pattern_syntax/case_pattern_empty_string_literal_match
# An empty string pattern "" in a case statement matches an empty target word precisely.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
m1=""; m2=""
case "" in
    "") m1="matched_empty" ;;
    *) m1="miss" ;;
esac
case "non_empty" in
    "") m2="matched_empty" ;;
    *) m2="miss" ;;
esac
[ "$m1" = "matched_empty" ] || fail "empty pattern should match empty target"
[ "$m2" = "miss" ] || fail "empty pattern should NOT match 'non_empty'"
echo PASS
exit 0
