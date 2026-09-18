#!/usr/bin/env bash
# vybe-test: bash/case_pattern_syntax/case_pattern_posix_character_class_upper
# POSIX character classes like [[:upper:]] are recognized within bracket patterns in case statements.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
m1=""; m2=""
case "G" in
    [[:upper:]]) m1="matched_upper" ;;
    *) m1="miss" ;;
esac
case "g" in
    [[:upper:]]) m2="matched_upper" ;;
    *) m2="miss" ;;
esac
[ "$m1" = "matched_upper" ] || fail "'G' should match [[:upper:]]"
[ "$m2" = "miss" ] || fail "'g' should not match [[:upper:]]"
echo PASS
exit 0
