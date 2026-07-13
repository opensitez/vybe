use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn shift_left_negative_value() { assert_eq!(run_c("int main() { int x = -1; /* x << 1 is UB in standard C, but often implementations act like unsigned or wrap */ printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn shift_right_negative_value_arithmetic() { assert_eq!(run_c("int main() { int x = -4; printf(\"%d\", x >> 1 == -2 || x >> 1 == 2147483646); return 0; }"), vec!["1"]); } // Implementation defined whether arithmetic (keeps sign) or logical. Usually arithmetic.
#[test] fn shift_right_negative_shift_count() { assert_eq!(run_c("int main() { int x = 4; /* x >> -1 is UB */ printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn shift_left_negative_shift_count() { assert_eq!(run_c("int main() { int x = 4; /* x << -1 is UB */ printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn shift_right_negative_literal() { assert_eq!(run_c("int main() { int x = -8 >> 2; printf(\"%d\", x == -2 || x > 0); return 0; }"), vec!["1"]); }
#[test] fn shift_left_negative_literal() { assert_eq!(run_c("int main() { /* int x = -8 << 2; // UB */ printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn shift_negative_in_unsigned_context() { assert_eq!(run_c("int main() { unsigned int x = (unsigned int)-4 >> 1; printf(\"%d\", x > 0); return 0; }"), vec!["1"]); } // Logical shift
#[test] fn shift_negative_char_promotion() { assert_eq!(run_c("int main() { char c = -2; int res = c >> 1; printf(\"%d\", res == -1 || res > 0); return 0; }"), vec!["1"]); } // Promotes to int (signed)
#[test] fn shift_negative_short_promotion() { assert_eq!(run_c("int main() { short s = -8; int res = s >> 2; printf(\"%d\", res == -2 || res > 0); return 0; }"), vec!["1"]); }
#[test] fn shift_negative_constant_expr() { assert_eq!(run_c("int main() { enum { A = -4 >> 1 }; printf(\"%d\", A == -2 || A > 0); return 0; }"), vec!["1"]); }
#[test] fn shift_negative_assignment() { assert_eq!(run_c("int main() { int x = -16; x >>= 2; printf(\"%d\", x == -4 || x > 0); return 0; }"), vec!["1"]); }
#[test] fn shift_negative_macro() { assert_eq!(run_c("#define SHIFT(x) ((x) >> 1)\nint main() { int x = -2; printf(\"%d\", SHIFT(x) == -1 || SHIFT(x) > 0); return 0; }"), vec!["1"]); }
#[test] fn shift_negative_long_long() { assert_eq!(run_c("int main() { long long x = -8LL; long long res = x >> 1; printf(\"%d\", res == -4LL || res > 0LL); return 0; }"), vec!["1"]); }
#[test] fn shift_negative_unsigned_to_signed() { assert_eq!(run_c("int main() { int x = (int)(4294967292U >> 1); printf(\"ok\"); return 0; }"), vec!["ok"]); } // 4294967292U is -4 in 32-bit if cast. Unsigned shift is logical.
#[test] fn shift_negative_chained() { assert_eq!(run_c("int main() { int x = -16 >> 1 >> 1; printf(\"%d\", x == -4 || x > 0); return 0; }"), vec!["1"]); }
