#!/usr/bin/env bash
# vybe-test: bash/case_pattern_syntax/case_pattern_extglob_exact_one_at
# In a case pattern, the extglob @(pattern1|pattern2) matches exactly one of the listed alternatives.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s extglob
m1=""; m2=""; m3=""
case "png" in
    @(jpg|png|gif)) m1="matched_png" ;;
    *) m1="miss" ;;
esac
case "pdf" in
    @(jpg|png|gif)) m2="matched_pdf" ;;
    *) m2="miss" ;;
esac
[ "$m1" = "matched_png" ] || fail "'png' should match @(jpg|png|gif)"
[ "$m2" = "miss" ] || fail "'pdf' should not match @(jpg|png|gif)"
echo PASS
exit 0
