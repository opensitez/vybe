#!/usr/bin/env bash
# vybe-test: bash/case_pattern_syntax/case_pattern_extglob_one_or_more_plus
# In a case pattern, the extglob +(pattern) construct matches one or more repetitions.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s extglob
matched=""
case "aaabbb" in
    +(a)+(b))
        matched="one_or_more"
        ;;
    *)
        matched="miss"
        ;;
esac
[ "$matched" = "one_or_more" ] || fail "+(a)+(b) extglob match failed: got [$matched]"
echo PASS
exit 0
