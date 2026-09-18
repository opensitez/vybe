#!/usr/bin/env bash
# vybe-test: bash/case_pattern_syntax/case_pattern_clause_test_next_double_semicolon_ampersand
# The ';;&' operator resumes testing subsequent patterns after executing the matching clause body.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
matches=""
case "alphabet" in
    alpha*)
        matches+="[prefix]"
        ;;&
    *bet)
        matches+="[suffix]"
        ;;&
    nomatch)
        matches+="[wrong]"
        ;;
esac
[ "$matches" = "[prefix][suffix]" ] || fail ";;& pattern testing failed: got [$matches]"
echo PASS
exit 0
