#!/usr/bin/env bash
# vybe-test: bash/parameter_expansion_basics/name_ends_at_first_non_name_character
# A name is letters, digits and underscore; anything else ends it and is
# copied literally.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=v
out=$(echo $x-1 $x.y $x/z $x@ $x:$x)
[ "$out" = 'v-1 v.y v/z v@ v:v' ] || fail "got [$out]"
echo PASS
exit 0
