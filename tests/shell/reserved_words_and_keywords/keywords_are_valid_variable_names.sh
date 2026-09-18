#!/usr/bin/env bash
# vybe-test: bash/reserved_words_and_keywords/keywords_are_valid_variable_names
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if=1; for=2; done=3; esac=4
sum=$((if + for + done + esac))
[ "$sum" = 10 ] || fail "want 10 got $sum"
[ "$if$for" = 12 ] || fail "want 12 got [$if$for]"
echo PASS
exit 0
