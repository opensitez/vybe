#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_width_and_precision/parent
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
width=$((10 % 6 + 2))
out=$(printf "%0${width}d" "10")
if ((${#out} < width)); then
  fail "width too small"
fi
prec=$((10 % 3 + 1))
out2=$(printf "%.${prec}f" "1.$((10 + 2))")
if ((${#out2} < prec + 2)); then
  fail "precision not respected"
fi
echo PASS
exit 0
