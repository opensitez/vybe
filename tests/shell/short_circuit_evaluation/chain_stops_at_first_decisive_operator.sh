#!/usr/bin/env bash
# vybe-test: bash/short_circuit_evaluation/chain_stops_at_first_decisive_operator
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(false || false || echo third && echo fourth || echo fifth)
[ "$out" = $'third\nfourth' ] || fail "got [$out]"
out=$(true || echo a || echo b && echo c)
[ "$out" = c ] || fail "got [$out]"
echo PASS
exit 0
