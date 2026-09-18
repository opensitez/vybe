#!/usr/bin/env bash
# vybe-test: bash/parameter_prefix_suffix_removal/removal_applies_to_each_positional
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- a.txt b.txt c.md
out="${@%.txt}"
[ "$out" = "a b c.md" ] || fail "suffix on \$@: got [$out]"
out="${*#?}"
[ "$out" = ".txt .txt .md" ] || fail "prefix on \$*: got [$out]"
count() { echo $#; }
[ "$(count "${@%.txt}")" = 3 ] || fail "word count must be preserved"
echo PASS
exit 0
