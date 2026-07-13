use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn math_fmod_positive() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", fmod(5.3, 2.0)); return 0; }"), vec!["1.3"]); }
#[test] fn math_fmod_negative() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", fmod(-5.3, 2.0)); return 0; }"), vec!["-1.3"]); }
#[test] fn math_fmod_negative_divisor() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", fmod(5.3, -2.0)); return 0; }"), vec!["1.3"]); }
#[test] fn math_fmod_exact() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", fmod(4.0, 2.0)); return 0; }"), vec!["0.0"]); }
#[test] fn math_remainder_positive() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", remainder(5.0, 3.0)); return 0; }"), vec!["-1.0"]); } // 5 - 3*2 = -1
#[test] fn math_remainder_halfway() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", remainder(5.0, 2.0)); return 0; }"), vec!["1.0"]); } // 5 / 2 = 2.5, nearest even is 2, 5 - 2*2 = 1.0
#[test] fn math_remainder_negative() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", remainder(-5.0, 3.0)); return 0; }"), vec!["1.0"]); } // -5 / 3 = -1.66 -> -2. -5 - (-2)*3 = 1
#[test] fn math_remquo_basic() { assert_eq!(run_c("#include <math.h>\nint main() { int q; double r = remquo(5.0, 3.0, &q); printf(\"%.1f %d\", r, q); return 0; }"), vec!["-1.0 2"]); }
#[test] fn math_remquo_negative() { assert_eq!(run_c("#include <math.h>\nint main() { int q; double r = remquo(-5.0, 3.0, &q); printf(\"%.1f %d\", r, q); return 0; }"), vec!["1.0 -2"]); }
#[test] fn math_fmod_div_zero() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%d\", isnan(fmod(5.0, 0.0))); return 0; }"), vec!["1"]); }
#[test] fn math_remainder_div_zero() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%d\", isnan(remainder(5.0, 0.0))); return 0; }"), vec!["1"]); }
#[test] fn math_fmod_inf() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%d\", isnan(fmod(INFINITY, 2.0))); return 0; }"), vec!["1"]); }
#[test] fn math_fmod_inf_divisor() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", fmod(5.0, INFINITY)); return 0; }"), vec!["5.0"]); }
#[test] fn math_remainder_inf_divisor() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", remainder(5.0, INFINITY)); return 0; }"), vec!["5.0"]); }
