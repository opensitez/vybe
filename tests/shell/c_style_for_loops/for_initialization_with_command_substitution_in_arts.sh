#!/usr/bin/env bash
# vybe-test: bash/c_style_for_loops/for_initialization_with_command_substitution_in_arts
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
i=($(printf 2))
for ((n=1;n<=i;n++)); do
  :
done
[ "$n" -eq 3 ] || fail "n=$n"
echo PASS
exit 0
