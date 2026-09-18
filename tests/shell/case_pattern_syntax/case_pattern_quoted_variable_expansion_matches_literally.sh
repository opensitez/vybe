#!/usr/bin/env bash
# vybe-test: bash/case_pattern_syntax/case_pattern_quoted_variable_expansion_matches_literally
# A quoted variable in a case pattern "$pat" is matched as a literal string rather than a glob.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
pat="star*"
m1=""; m2=""
case "star*" in
    "$pat") m1="matched_literal" ;;
    *) m1="miss" ;;
esac
case "star_extended" in
    "$pat") m2="matched_wildcard" ;;
    *) m2="miss" ;;
esac
[ "$m1" = "matched_literal" ] || fail "quoted variable should match literal string"
[ "$m2" = "miss" ] || fail "quoted variable should NOT expand wildcard"
echo PASS
exit 0
