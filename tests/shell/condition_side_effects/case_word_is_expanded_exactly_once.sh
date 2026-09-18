#!/usr/bin/env bash
# vybe-test: bash/condition_side_effects/case_word_is_expanded_exactly_once
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
i=0
case $((i++)) in
  5) r=five ;;
  0) r=zero ;;
  *) r=other ;;
esac
[ "$r" = zero ] || fail "matched [$r]"
[ "$i" -eq 1 ] || fail "word expanded once, i=$i"
echo PASS
exit 0
