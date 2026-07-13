use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn math_asin_zero() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", asin(0.0)); return 0; }"), vec!["0.0"]); }
#[test] fn math_acos_one() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", acos(1.0)); return 0; }"), vec!["0.0"]); }
#[test] fn math_atan_zero() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", atan(0.0)); return 0; }"), vec!["0.0"]); }
#[test] fn math_asin_one() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.5f\", asin(1.0)); return 0; }"), vec!["1.57080"]); } // Pi/2
#[test] fn math_acos_zero() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.5f\", acos(0.0)); return 0; }"), vec!["1.57080"]); } // Pi/2
#[test] fn math_atan_one() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.5f\", atan(1.0)); return 0; }"), vec!["0.78540"]); } // Pi/4
#[test] fn math_asin_negative_one() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.5f\", asin(-1.0)); return 0; }"), vec!["-1.57080"]); }
#[test] fn math_acos_negative_one() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.5f\", acos(-1.0)); return 0; }"), vec!["3.14159"]); } // Pi
#[test] fn math_atan_negative_one() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.5f\", atan(-1.0)); return 0; }"), vec!["-0.78540"]); }
#[test] fn math_atan2_first_quadrant() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.5f\", atan2(1.0, 1.0)); return 0; }"), vec!["0.78540"]); }
#[test] fn math_atan2_second_quadrant() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.5f\", atan2(1.0, -1.0)); return 0; }"), vec!["2.35619"]); } // 3Pi/4
#[test] fn math_atan2_third_quadrant() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.5f\", atan2(-1.0, -1.0)); return 0; }"), vec!["-2.35619"]); } // -3Pi/4
#[test] fn math_atan2_fourth_quadrant() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.5f\", atan2(-1.0, 1.0)); return 0; }"), vec!["-0.78540"]); } // -Pi/4
#[test] fn math_asin_out_of_bounds() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%d\", isnan(asin(2.0))); return 0; }"), vec!["1"]); } // Domain error
