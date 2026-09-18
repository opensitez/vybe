#!/usr/bin/env bash
# vybe-test: bash/case_pattern_syntax/case_pattern_extglob_zero_or_one_question
# In a case pattern, the extglob ?(pattern) construct matches zero or one occurrence.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s extglob
m0=""; m1=""; m2=""
case "colour" in
    col?(o)ur) m1="matched_one" ;;
    *) m1="miss" ;;
esac
case "colur" in
    col?(o)ur) m0="matched_zero" ;;
    *) m0="miss" ;;
esac
case "coloor" in
    col?(o)ur) m2="matched_two" ;;
    *) m2="miss" ;;
esac
[ "$m1" = "matched_one" ] || fail "one occurrence failed"
[ "$m0" = "matched_zero" ] || fail "zero occurrence failed"
[ "$m2" = "miss" ] || fail "two occurrences should not match"
echo PASS
exit 0
