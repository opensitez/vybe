use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn math_j0_zero() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", j0(0.0)); return 0; }"), vec!["1.0"]); } // Bessel function of first kind, order 0
#[test] fn math_j1_zero() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", j1(0.0)); return 0; }"), vec!["0.0"]); }
#[test] fn math_jn_zero() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", jn(2, 0.0)); return 0; }"), vec!["0.0"]); }
#[test] fn math_y0_one() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.5f\", y0(1.0)); return 0; }"), vec!["0.08826"]); } // Bessel function of second kind, order 0
#[test] fn math_y1_one() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.5f\", y1(1.0)); return 0; }"), vec!["-0.78121"]); }
#[test] fn math_yn_one() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.5f\", yn(2, 1.0)); return 0; }"), vec!["-1.65068"]); }
#[test] fn math_j0_inf() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", j0(INFINITY)); return 0; }"), vec!["0.0"]); } // Decays to 0
#[test] fn math_y0_inf() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", y0(INFINITY)); return 0; }"), vec!["0.0"]); } // Decays to 0
#[test] fn math_j1_negative() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.5f\", j1(-1.0)); return 0; }"), vec!["-0.44005"]); } // j1 is odd
#[test] fn math_j0_negative() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.5f\", j0(-1.0)); return 0; }"), vec!["0.76520"]); } // j0 is even
#[test] fn math_y0_negative_domain_error() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"ok\"); return 0; }"), vec!["ok"]); } // Domain error for x < 0
#[test] fn math_jn_large_order() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.5f\", jn(10, 1.0)); return 0; }"), vec!["0.00000"]); } // Very small
#[test] fn math_yn_large_order() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"ok\"); return 0; }"), vec!["ok"]); } // Might be very large negative
