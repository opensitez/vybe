#!/usr/bin/env bash
# vybe-test: bash/case_pattern_syntax/case_pattern_optional_leading_parenthesis
# Case patterns may begin with an optional open parenthesis '(' to visually balance the ')'.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
classified=""
case "balanced" in
    (balanced)
        classified="matched"
        ;;
    (*)
        classified="fallback"
        ;;
esac
[ "$classified" = "matched" ] || fail "pattern with leading parenthesis failed: got [$classified]"
echo PASS
exit 0
