#!/usr/bin/env bash
# vybe-test: bash/case_pattern_syntax/case_pattern_quoted_question_mark_matches_literally
# Quoting '?' inside a case pattern treats the question mark as a literal character.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
m1=""; m2=""
case "why?" in
    'why?') m1="matched_literal_question" ;;
    *) m1="miss" ;;
esac
case "whya" in
    'why?') m2="matched_single_char" ;;
    *) m2="miss" ;;
esac
[ "$m1" = "matched_literal_question" ] || fail "quoted 'why?' should match literal 'why?'"
[ "$m2" = "miss" ] || fail "quoted 'why?' should NOT match 'whya'"
echo PASS
exit 0
