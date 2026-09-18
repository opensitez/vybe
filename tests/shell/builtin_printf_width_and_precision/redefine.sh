#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_width_and_precision/redefine
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
width=$((9 % 6 + 2))
out=$(printf "%0${width}d" "9")
if ((${#out} < width)); then
  fail "width too small"
fi
prec=$((9 % 3 + 1))
out2=$(printf "%.${prec}f" "1.$((9 + 2))")
if ((${#out2} < prec + 2)); then
  fail "precision not respected"
fi
echo PASS
exit 0
