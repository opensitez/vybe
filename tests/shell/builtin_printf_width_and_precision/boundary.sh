#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_width_and_precision/boundary
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
width=$((13 % 6 + 2))
out=$(printf "%0${width}d" "13")
if ((${#out} < width)); then
  fail "width too small"
fi
prec=$((13 % 3 + 1))
out2=$(printf "%.${prec}f" "1.$((13 + 2))")
if ((${#out2} < prec + 2)); then
  fail "precision not respected"
fi
echo PASS
exit 0
