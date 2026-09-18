#!/usr/bin/env bash
# vybe-test: bash/case_pattern_syntax/case_pattern_character_class_caret_inversion
# A caret at the start of a character class [^... ] also inverts the character set in Bash.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
m1=""; m2=""
case "!" in
    [^0-9]) m1="non_digit" ;;
    *) m1="digit" ;;
esac
case "9" in
    [^0-9]) m2="non_digit" ;;
    *) m2="digit" ;;
esac
[ "$m1" = "non_digit" ] || fail "'!' should match [^0-9]"
[ "$m2" = "digit" ] || fail "'9' should not match [^0-9]"
echo PASS
exit 0
