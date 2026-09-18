#!/usr/bin/env bash
# vybe-test: bash/case_pattern_syntax/case_pattern_extglob_zero_or_more_star
# In a case pattern, the extglob *(pattern) construct matches zero or more occurrences.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s extglob
m1=""; m2=""
case "prefix" in
    prefix*(ext)) m1="matched_zero" ;;
    *) m1="miss" ;;
esac
case "prefixextext" in
    prefix*(ext)) m2="matched_two" ;;
    *) m2="miss" ;;
esac
[ "$m1" = "matched_zero" ] || fail "zero occurrences failed"
[ "$m2" = "matched_two" ] || fail "multiple occurrences failed"
echo PASS
exit 0
