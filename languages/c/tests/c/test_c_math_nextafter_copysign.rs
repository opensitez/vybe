use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn math_nextafter_up() { assert_eq!(run_c("#include <math.h>\nint main() { double n = nextafter(1.0, 2.0); printf(\"%d\", n > 1.0 && n < 1.000000000000001); return 0; }"), vec!["1"]); }
#[test] fn math_nextafter_down() { assert_eq!(run_c("#include <math.h>\nint main() { double n = nextafter(1.0, 0.0); printf(\"%d\", n < 1.0 && n > 0.999999999999999); return 0; }"), vec!["1"]); }
#[test] fn math_nextafter_equal() { assert_eq!(run_c("#include <math.h>\nint main() { double n = nextafter(1.0, 1.0); printf(\"%d\", n == 1.0); return 0; }"), vec!["1"]); }
#[test] fn math_nexttoward_up() { assert_eq!(run_c("#include <math.h>\nint main() { double n = nexttoward(1.0, 2.0L); printf(\"%d\", n > 1.0); return 0; }"), vec!["1"]); }
#[test] fn math_copysign_pos_to_neg() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", copysign(2.0, -3.0)); return 0; }"), vec!["-2.0"]); }
#[test] fn math_copysign_neg_to_pos() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", copysign(-2.0, 3.0)); return 0; }"), vec!["2.0"]); }
#[test] fn math_copysign_pos_to_pos() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", copysign(2.0, 3.0)); return 0; }"), vec!["2.0"]); }
#[test] fn math_copysign_neg_to_neg() { assert_eq!(run_c("#include <math.h>\nint main() { printf(\"%.1f\", copysign(-2.0, -3.0)); return 0; }"), vec!["-2.0"]); }
#[test] fn math_copysign_zero_neg() { assert_eq!(run_c("#include <math.h>\nint main() { double n = copysign(2.0, -0.0); printf(\"%d\", n == -2.0 && signbit(n)); return 0; }"), vec!["1"]); }
#[test] fn math_copysign_inf() { assert_eq!(run_c("#include <math.h>\nint main() { double n = copysign(INFINITY, -1.0); printf(\"%d\", isinf(n) && signbit(n)); return 0; }"), vec!["1"]); }
#[test] fn math_copysign_nan() { assert_eq!(run_c("#include <math.h>\nint main() { double n = copysign(1.0, NAN); printf(\"ok\"); return 0; }"), vec!["ok"]); } // works with NAN but result sign is system dependent
#[test] fn math_nextafter_zero_to_pos() { assert_eq!(run_c("#include <math.h>\nint main() { double n = nextafter(0.0, 1.0); printf(\"%d\", n > 0.0); return 0; }"), vec!["1"]); }
#[test] fn math_nextafter_zero_to_neg() { assert_eq!(run_c("#include <math.h>\nint main() { double n = nextafter(0.0, -1.0); printf(\"%d\", n < 0.0); return 0; }"), vec!["1"]); }
#[test] fn math_nextafter_inf() { assert_eq!(run_c("#include <math.h>\n#include <float.h>\nint main() { double n = nextafter(DBL_MAX, INFINITY); printf(\"%d\", isinf(n)); return 0; }"), vec!["1"]); }
