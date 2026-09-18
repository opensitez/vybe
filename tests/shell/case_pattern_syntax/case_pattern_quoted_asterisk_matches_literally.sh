#!/usr/bin/env bash
# vybe-test: bash/case_pattern_syntax/case_pattern_quoted_asterisk_matches_literally
# Quoting '*' inside a case pattern treats the asterisk as a literal character, not a wildcard.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
m1=""; m2=""
case "star*" in
    "star*") m1="matched_literal" ;;
    *) m1="miss" ;;
esac
case "star_other" in
    "star*") m2="matched_wildcard" ;;
    *) m2="miss" ;;
esac
[ "$m1" = "matched_literal" ] || fail "quoted asterisk should match literal 'star*'"
[ "$m2" = "miss" ] || fail "quoted asterisk should NOT match 'star_other'"
echo PASS
exit 0
