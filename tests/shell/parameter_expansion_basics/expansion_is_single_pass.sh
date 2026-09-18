#!/usr/bin/env bash
# vybe-test: bash/parameter_expansion_basics/expansion_is_single_pass
# The text produced by an expansion is not scanned for further $ expansions.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
y=inner
x='$y'
out=$(echo $x "$x")
[ "$out" = '$y $y' ] || fail "want [\$y \$y] got [$out]"
echo PASS
exit 0
