#!/usr/bin/env bash
# vybe-test: bash/autoload_function_patterns/version_probe
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=18
prefix="autoload_${IDX}_"
eval "${prefix}one(){ :; }"
eval "${prefix}two(){ :; }"
all=$(declare -F "${prefix}*")
[[ $all == *"${prefix}one"* ]] || fail "wildcard listing missing first function"
[[ $all == *"${prefix}two"* ]] || fail "wildcard listing missing second function"
single=$(declare -F "${prefix}one")
[[ $single == *"${prefix}one"* ]] || fail "exact listing missing first function"
if (( IDX % 2 == 0 )); then
  missing=$(declare -F "${prefix}ghost*" 2>/dev/null || true)
  [[ -z $missing ]] || fail "unexpected missing-match declaration output"
fi
if (( IDX % 3 == 0 )); then
  maybe=$(declare -F "${prefix}o?e" 2>/dev/null || true)
  [[ -n $maybe ]] || fail "question-pattern should match known function"
fi
if (( IDX % 5 == 0 )); then
  eval "${prefix}third(){ return 0; }"
  ${prefix}third
  (( $? == 0 )) || fail "dynamically created function did not run"
fi
unset -f "${prefix}one" "${prefix}two" 2>/dev/null || true
unset -f "${prefix}third" 2>/dev/null || true
echo PASS
exit 0
