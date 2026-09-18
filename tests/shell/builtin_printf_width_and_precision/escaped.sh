#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_width_and_precision/escaped
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
width=$((7 % 6 + 2))
out=$(printf "%0${width}d" "7")
if ((${#out} < width)); then
  fail "width too small"
fi
prec=$((7 % 3 + 1))
out2=$(printf "%.${prec}f" "1.$((7 + 2))")
if ((${#out2} < prec + 2)); then
  fail "precision not respected"
fi
echo PASS
exit 0
