#!/usr/bin/env bash
# vybe-test: bash/case_pattern_syntax/case_pattern_character_class_range
# A bracketed character range [a-z] matches any single character within the range.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
m1=""; m2=""
case "m" in
    [a-z]) m1="lowercase_matched" ;;
    *) m1="miss" ;;
esac
case "5" in
    [0-9]) m2="digit_matched" ;;
    *) m2="miss" ;;
esac
[ "$m1" = "lowercase_matched" ] || fail "character class [a-z] failed"
[ "$m2" = "digit_matched" ] || fail "character class [0-9] failed"
echo PASS
exit 0
