#!/usr/bin/env bash
# vybe-test: bash/expansion_order_and_interactions/variables_inside_brace_elements_expand_after_brace
# Brace expansion produces the words first; parameter expansion then runs on
# each result. So {$a,$b}x becomes $ax and $bx (two other variables), while
# {${a},${b}}x keeps the names intact.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=1; b=2; unset ax bx
out=$(echo {${a},${b}}x)
[ "$out" = "1x 2x" ] || fail "braced names: got [$out]"
out=$(echo {$a,$b}x)
[ -z "$out" ] || fail "\$ax and \$bx are unset, want empty got [$out]"
echo PASS
exit 0
