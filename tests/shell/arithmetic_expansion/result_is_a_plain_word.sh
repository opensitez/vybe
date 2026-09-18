#!/usr/bin/env bash
# vybe-test: bash/arithmetic_expansion/result_is_a_plain_word
# The result joins adjacent text into one word and is subject to word
# splitting only if it contained IFS characters, which a number never does.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
[ "$(echo item$((1+1))x)" = item2x ] || fail "got [$(echo item$((1+1))x)]"
[ "$(count $((10-3)) $((2*2)))" = 2 ] || fail "two results, two words"
echo PASS
exit 0
