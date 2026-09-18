#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_width_and_precision/override
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
width=$((4 % 6 + 2))
out=$(printf "%0${width}d" "4")
if ((${#out} < width)); then
  fail "width too small"
fi
prec=$((4 % 3 + 1))
out2=$(printf "%.${prec}f" "1.$((4 + 2))")
if ((${#out2} < prec + 2)); then
  fail "precision not respected"
fi
echo PASS
exit 0
