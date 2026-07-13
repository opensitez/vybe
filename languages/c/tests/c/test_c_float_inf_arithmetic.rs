use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn float_inf_macro() { assert_eq!(run_c("#include <math.h>\nint main() { double n = INFINITY; printf(\"%d\", isinf(n)); return 0; }"), vec!["1"]); }
#[test] fn float_inf_addition() { assert_eq!(run_c("#include <math.h>\nint main() { double n = INFINITY + 5.0; printf(\"%d\", isinf(n) && n > 0); return 0; }"), vec!["1"]); }
#[test] fn float_inf_subtraction() { assert_eq!(run_c("#include <math.h>\nint main() { double n = INFINITY - 1000.0; printf(\"%d\", isinf(n) && n > 0); return 0; }"), vec!["1"]); }
#[test] fn float_inf_multiplication() { assert_eq!(run_c("#include <math.h>\nint main() { double n = INFINITY * 2.0; printf(\"%d\", isinf(n) && n > 0); return 0; }"), vec!["1"]); }
#[test] fn float_inf_multiplication_negative() { assert_eq!(run_c("#include <math.h>\nint main() { double n = INFINITY * -2.0; printf(\"%d\", isinf(n) && n < 0); return 0; }"), vec!["1"]); }
#[test] fn float_inf_division() { assert_eq!(run_c("#include <math.h>\nint main() { double n = INFINITY / 2.0; printf(\"%d\", isinf(n) && n > 0); return 0; }"), vec!["1"]); }
#[test] fn float_inf_div_by_inf_is_nan() { assert_eq!(run_c("#include <math.h>\nint main() { double n = INFINITY / INFINITY; printf(\"%d\", isnan(n)); return 0; }"), vec!["1"]); }
#[test] fn float_inf_minus_inf_is_nan() { assert_eq!(run_c("#include <math.h>\nint main() { double n = INFINITY - INFINITY; printf(\"%d\", isnan(n)); return 0; }"), vec!["1"]); }
#[test] fn float_inf_plus_inf() { assert_eq!(run_c("#include <math.h>\nint main() { double n = INFINITY + INFINITY; printf(\"%d\", isinf(n)); return 0; }"), vec!["1"]); }
#[test] fn float_div_zero_is_inf() { assert_eq!(run_c("#include <math.h>\nint main() { double n = 1.0 / 0.0; printf(\"%d\", isinf(n) && n > 0); return 0; }"), vec!["1"]); }
#[test] fn float_div_neg_zero_is_neg_inf() { assert_eq!(run_c("#include <math.h>\nint main() { double n = 1.0 / -0.0; printf(\"%d\", isinf(n) && n < 0); return 0; }"), vec!["1"]); }
#[test] fn float_inf_comparison() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%d\", INFINITY > 1e100); return 0; }"), vec!["1"]); }
#[test] fn float_inf_equality() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%d\", INFINITY == INFINITY); return 0; }"), vec!["1"]); }
#[test] fn float_inf_pow() { assert_eq!(run_c("#include <math.h>\nint main() { double n = pow(2.0, INFINITY); printf(\"%d\", isinf(n)); return 0; }"), vec!["1"]); }
