#!/usr/bin/env bash
# vybe-test: bash/case_pattern_syntax/case_pattern_wildcard_asterisk_matches_any
# The '*' pattern wildcard matches any string, including empty strings.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
m1=""; m2=""
case "prefix_123" in
    prefix_*) m1="matched_prefix" ;;
    *) m1="miss" ;;
esac
case "prefix_" in
    prefix_*) m2="matched_empty_tail" ;;
    *) m2="miss" ;;
esac
[ "$m1" = "matched_prefix" ] || fail "m1 failed"
[ "$m2" = "matched_empty_tail" ] || fail "m2 failed"
echo PASS
exit 0
