#!/usr/bin/env bash
# vybe-test: bash/arithmetic_increment_and_assignment/compound_assignment_operators
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=10
out="$((x+=5)) $((x-=3)) $((x*=2)) $((x/=4)) $((x%=4)) $((x<<=3)) $((x>>=1)) $((x&=6)) $((x|=1)) $((x^=3))"
[ "$out" = "15 12 24 6 2 16 8 0 1 2" ] || fail "got [$out]"
[ "$x" -eq 2 ] || fail "final x want 2 got $x"
echo PASS
exit 0
