#!/usr/bin/env bash
# vybe-test: bash/case_pattern_syntax/case_pattern_literal_string_exact_match
# An unquoted alphanumeric pattern matches an identical string literal.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
match=""
case "target_word" in
    target_word)
        match="exact"
        ;;
    *)
        match="fallback"
        ;;
esac
[ "$match" = "exact" ] || fail "literal match failed: got [$match]"
echo PASS
exit 0
