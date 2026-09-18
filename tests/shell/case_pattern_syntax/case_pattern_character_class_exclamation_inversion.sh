#!/usr/bin/env bash
# vybe-test: bash/case_pattern_syntax/case_pattern_character_class_exclamation_inversion
# An exclamation point at the start of a character class [!... ] inverts the pattern match.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
m1=""; m2=""
case "7" in
    [!a-zA-Z]) m1="non_alpha" ;;
    *) m1="alpha" ;;
esac
case "A" in
    [!a-zA-Z]) m2="non_alpha" ;;
    *) m2="alpha" ;;
esac
[ "$m1" = "non_alpha" ] || fail "'7' should match [!a-zA-Z]"
[ "$m2" = "alpha" ] || fail "'A' should not match [!a-zA-Z]"
echo PASS
exit 0
