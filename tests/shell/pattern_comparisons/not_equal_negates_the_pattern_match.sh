#!/usr/bin/env bash
# vybe-test: bash/pattern_comparisons/not_equal_negates_the_pattern_match
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ abc != a* ]] && fail "abc matches a*, so != must be false"
[[ xyz != a* ]] || fail "xyz does not match a*, so != must be true"
echo PASS
exit 0
