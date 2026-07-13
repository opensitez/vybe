use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn math_sinh_zero() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", sinh(0.0)); return 0; }"), vec!["0.0"]); }
#[test] fn math_cosh_zero() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", cosh(0.0)); return 0; }"), vec!["1.0"]); }
#[test] fn math_tanh_zero() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", tanh(0.0)); return 0; }"), vec!["0.0"]); }
#[test] fn math_sinh_one() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.5f\", sinh(1.0)); return 0; }"), vec!["1.17520"]); }
#[test] fn math_cosh_one() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.5f\", cosh(1.0)); return 0; }"), vec!["1.54308"]); }
#[test] fn math_tanh_one() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.5f\", tanh(1.0)); return 0; }"), vec!["0.76159"]); }
#[test] fn math_sinh_negative() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.5f\", sinh(-1.0)); return 0; }"), vec!["-1.17520"]); } // sinh is odd
#[test] fn math_cosh_negative() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.5f\", cosh(-1.0)); return 0; }"), vec!["1.54308"]); } // cosh is even
#[test] fn math_tanh_negative() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.5f\", tanh(-1.0)); return 0; }"), vec!["-0.76159"]); } // tanh is odd
#[test] fn math_asinh_zero() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", asinh(0.0)); return 0; }"), vec!["0.0"]); }
#[test] fn math_acosh_one() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", acosh(1.0)); return 0; }"), vec!["0.0"]); }
#[test] fn math_atanh_zero() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", atanh(0.0)); return 0; }"), vec!["0.0"]); }
#[test] fn math_acosh_out_of_bounds() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%d\", isnan(acosh(0.5))); return 0; }"), vec!["1"]); } // x must be >= 1
#[test] fn math_atanh_out_of_bounds() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%d\", isnan(atanh(2.0))); return 0; }"), vec!["1"]); } // x must be in [-1, 1]
