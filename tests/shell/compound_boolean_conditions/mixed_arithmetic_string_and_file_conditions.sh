#!/usr/bin/env bash
# vybe-test: bash/compound_boolean_conditions/mixed_arithmetic_string_and_file_conditions
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
n=5; s=alpha
if (( n > 1 )) && [[ $s == a* ]] && [ -d / ]; then r=yes; else r=no; fi
[ "$r" = yes ] || fail "all three true: got $r"
if (( n > 10 )) || [[ $s == b* ]] || [ -d /nonexistent_zz ]; then r=yes; else r=no; fi
[ "$r" = no ] || fail "all three false: got $r"
echo PASS
exit 0
