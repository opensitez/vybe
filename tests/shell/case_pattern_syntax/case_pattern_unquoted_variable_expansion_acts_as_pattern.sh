#!/usr/bin/env bash
# vybe-test: bash/case_pattern_syntax/case_pattern_unquoted_variable_expansion_acts_as_pattern
# An unquoted variable in a case pattern has its value expanded and interpreted as a pattern.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
pat="val_*"
matched=""
case "val_12345" in
    $pat)
        matched="expanded_wildcard"
        ;;
    *)
        matched="miss"
        ;;
esac
[ "$matched" = "expanded_wildcard" ] || fail "unquoted variable pattern failed: got [$matched]"
echo PASS
exit 0
