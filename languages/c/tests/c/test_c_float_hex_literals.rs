use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn float_hex_basic() { assert_eq!(run_c("int main() { printf(\"%f\", 0x1.0p0); return 0; }"), vec!["1.000000"]); }
#[test] fn float_hex_fraction() { assert_eq!(run_c("int main() { printf(\"%f\", 0x1.8p0); return 0; }"), vec!["1.500000"]); } // 1 + 8/16 = 1.5
#[test] fn float_hex_positive_exponent() { assert_eq!(run_c("int main() { printf(\"%f\", 0x1.0p4); return 0; }"), vec!["16.000000"]); } // 1 * 2^4 = 16
#[test] fn float_hex_negative_exponent() { assert_eq!(run_c("int main() { printf(\"%f\", 0x1.0p-1); return 0; }"), vec!["0.500000"]); } // 1 * 2^-1 = 0.5
#[test] fn float_hex_no_fraction() { assert_eq!(run_c("int main() { printf(\"%f\", 0x2p3); return 0; }"), vec!["16.000000"]); } // 2 * 2^3 = 16
#[test] fn float_hex_large_mantissa() { assert_eq!(run_c("int main() { printf(\"%f\", 0xFF.0p-4); return 0; }"), vec!["15.937500"]); } // 255 / 16 = 15.9375
#[test] fn float_hex_float_suffix() { assert_eq!(run_c("int main() { printf(\"%f\", 0x1.0p0f); return 0; }"), vec!["1.000000"]); }
#[test] fn float_hex_long_double_suffix() { assert_eq!(run_c("int main() { printf(\"%f\", (double)0x1.0p0L); return 0; }"), vec!["1.000000"]); }
#[test] fn float_hex_zero() { assert_eq!(run_c("int main() { printf(\"%f\", 0x0.0p0); return 0; }"), vec!["0.000000"]); }
#[test] fn float_hex_negative() { assert_eq!(run_c("int main() { printf(\"%f\", -0x1.cp1); return 0; }"), vec!["-3.500000"]); } // -(1 + 12/16) * 2 = -1.75 * 2 = -3.5
#[test] fn float_hex_uppercase() { assert_eq!(run_c("int main() { printf(\"%f\", 0X1.AP2); return 0; }"), vec!["6.500000"]); } // (1 + 10/16) * 4 = 1.625 * 4 = 6.5
#[test] fn float_hex_printf_a() { assert_eq!(run_c("int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); } // %a formatting is implementation defined in exact form, let's just test compile
#[test] fn float_hex_scanf_a() { assert_eq!(run_c("int main() { float f; sscanf(\"0x1.8p1\", \"%a\", &f); printf(\"%f\", f); return 0; }"), vec!["3.000000"]); } // 1.5 * 2 = 3.0
#[test] fn float_hex_subnormal() { assert_eq!(run_c("int main() { double f = 0x0.000001p-1000; printf(\"%d\", f > 0.0); return 0; }"), vec!["1"]); } // valid double
