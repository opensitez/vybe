#!/usr/bin/env bash
# vybe-test: bash/process_substitution/quoted_substitution_is_literal_text
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(echo "<(true)" '>(true)')
[ "$out" = '<(true) >(true)' ] || fail "got [$out]"
echo PASS
exit 0
