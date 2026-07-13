use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn math_ceil_positive() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", ceil(2.3)); return 0; }"), vec!["3.0"]); }
#[test] fn math_ceil_negative() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", ceil(-2.3)); return 0; }"), vec!["-2.0"]); }
#[test] fn math_ceil_exact() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", ceil(2.0)); return 0; }"), vec!["2.0"]); }
#[test] fn math_floor_positive() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", floor(2.8)); return 0; }"), vec!["2.0"]); }
#[test] fn math_floor_negative() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", floor(-2.8)); return 0; }"), vec!["-3.0"]); }
#[test] fn math_floor_exact() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", floor(2.0)); return 0; }"), vec!["2.0"]); }
#[test] fn math_ceil_zero() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", ceil(0.0)); return 0; }"), vec!["0.0"]); }
#[test] fn math_floor_zero() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", floor(0.0)); return 0; }"), vec!["0.0"]); }
#[test] fn math_ceil_inf() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%d\", isinf(ceil(INFINITY))); return 0; }"), vec!["1"]); }
#[test] fn math_floor_inf() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%d\", isinf(floor(INFINITY))); return 0; }"), vec!["1"]); }
#[test] fn math_ceil_nan() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%d\", isnan(ceil(NAN))); return 0; }"), vec!["1"]); }
#[test] fn math_floor_nan() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%d\", isnan(floor(NAN))); return 0; }"), vec!["1"]); }
#[test] fn math_ceil_negative_zero() { assert_eq!(run_c("#include <math.h>\nint main() { double n = ceil(-0.0); printf(\"%d\", n == 0.0 && signbit(n)); return 0; }"), vec!["1"]); }
#[test] fn math_floor_negative_zero() { assert_eq!(run_c("#include <math.h>\nint main() { double n = floor(-0.0); printf(\"%d\", n == 0.0 && signbit(n)); return 0; }"), vec!["1"]); }
