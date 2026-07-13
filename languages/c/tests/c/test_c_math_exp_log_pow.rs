use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn math_exp_zero() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", exp(0.0)); return 0; }"), vec!["1.0"]); }
#[test] fn math_exp_one() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.5f\", exp(1.0)); return 0; }"), vec!["2.71828"]); } // e
#[test] fn math_exp2_basic() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", exp2(3.0)); return 0; }"), vec!["8.0"]); } // 2^3
#[test] fn math_expm1_zero() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", expm1(0.0)); return 0; }"), vec!["0.0"]); } // exp(0) - 1
#[test] fn math_log_e() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", log(exp(1.0))); return 0; }"), vec!["1.0"]); }
#[test] fn math_log10_basic() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", log10(100.0)); return 0; }"), vec!["2.0"]); }
#[test] fn math_log2_basic() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", log2(8.0)); return 0; }"), vec!["3.0"]); }
#[test] fn math_log1p_zero() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", log1p(0.0)); return 0; }"), vec!["0.0"]); } // log(1+0)
#[test] fn math_pow_basic() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", pow(2.0, 3.0)); return 0; }"), vec!["8.0"]); }
#[test] fn math_pow_zero_exponent() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", pow(5.0, 0.0)); return 0; }"), vec!["1.0"]); }
#[test] fn math_pow_negative_base_integer_exp() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", pow(-2.0, 3.0)); return 0; }"), vec!["-8.0"]); }
#[test] fn math_pow_negative_base_fractional_exp_nan() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%d\", isnan(pow(-2.0, 1.5))); return 0; }"), vec!["1"]); } // Domain error
#[test] fn math_sqrt_basic() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", sqrt(9.0)); return 0; }"), vec!["3.0"]); }
#[test] fn math_cbrt_basic() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", cbrt(27.0)); return 0; }"), vec!["3.0"]); }
#[test] fn math_cbrt_negative() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", cbrt(-8.0)); return 0; }"), vec!["-2.0"]); }
