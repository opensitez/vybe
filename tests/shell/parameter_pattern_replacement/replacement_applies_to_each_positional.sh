#!/usr/bin/env bash
# vybe-test: bash/parameter_pattern_replacement/replacement_applies_to_each_positional
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- cat hat bat
out="${@/a/o}"
[ "$out" = "cot hot bot" ] || fail "got [$out]"
count() { echo $#; }
[ "$(count "${@//t/}")" = 3 ] || fail "word count preserved"
echo PASS
exit 0
