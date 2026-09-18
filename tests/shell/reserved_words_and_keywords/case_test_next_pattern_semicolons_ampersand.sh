#!/usr/bin/env bash
# vybe-test: bash/reserved_words_and_keywords/case_test_next_pattern_semicolons_ampersand
# The ;;& operator in a case statement tests subsequent patterns after executing the matching clause body.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
acc=""
case "foobar" in
    foo*)
        acc+="prefix_"
        ;;&
    *bar)
        acc+="suffix"
        ;;
    nomatch)
        acc+="wrong"
        ;;
esac
[ "$acc" = "prefix_suffix" ] || fail "case resume testing with ;;&: want 'prefix_suffix', got [$acc]"
echo PASS
exit 0
