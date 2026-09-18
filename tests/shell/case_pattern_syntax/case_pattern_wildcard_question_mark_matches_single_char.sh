#!/usr/bin/env bash
# vybe-test: bash/case_pattern_syntax/case_pattern_wildcard_question_mark_matches_single_char
# The '?' pattern wildcard matches exactly one single character.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
m1=""; m2=""
case "cat" in
    c?t) m1="matched_three_chars" ;;
    *) m1="miss" ;;
esac
case "cart" in
    c?t) m2="matched_four_chars" ;;
    *) m2="miss" ;;
esac
[ "$m1" = "matched_three_chars" ] || fail "c?t should match 'cat'"
[ "$m2" = "miss" ] || fail "c?t should not match 'cart'"
echo PASS
exit 0
