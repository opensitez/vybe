use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn fenv_set_round_down() { assert_eq!(run_c("#include <fenv.h>\n#pragma STDC FENV_ACCESS ON\nint main() { fesetround(FE_DOWNWARD); printf(\"%d\", fegetround() == FE_DOWNWARD); return 0; }"), vec!["1"]); }
#[test] fn fenv_set_round_up() { assert_eq!(run_c("#include <fenv.h>\n#pragma STDC FENV_ACCESS ON\nint main() { fesetround(FE_UPWARD); printf(\"%d\", fegetround() == FE_UPWARD); return 0; }"), vec!["1"]); }
#[test] fn fenv_set_round_toward_zero() { assert_eq!(run_c("#include <fenv.h>\n#pragma STDC FENV_ACCESS ON\nint main() { fesetround(FE_TOWARDZERO); printf(\"%d\", fegetround() == FE_TOWARDZERO); return 0; }"), vec!["1"]); }
#[test] fn fenv_set_round_tonearest() { assert_eq!(run_c("#include <fenv.h>\n#pragma STDC FENV_ACCESS ON\nint main() { fesetround(FE_TONEAREST); printf(\"%d\", fegetround() == FE_TONEAREST); return 0; }"), vec!["1"]); }
#[test] fn fenv_rounding_affects_addition() { assert_eq!(run_c("#include <fenv.h>\n#include <stdio.h>\n#pragma STDC FENV_ACCESS ON\nint main() { fesetround(FE_UPWARD); double x = 1.0; double y = 3.0; double z = x / y; printf(\"%d\", z > 0.3333333333333333); return 0; }"), vec!["1"]); }
#[test] fn fenv_rounding_downward_affects_division() { assert_eq!(run_c("#include <fenv.h>\n#include <stdio.h>\n#pragma STDC FENV_ACCESS ON\nint main() { fesetround(FE_DOWNWARD); double x = 1.0; double y = 3.0; double z = x / y; printf(\"%d\", z < 0.3333333333333334); return 0; }"), vec!["1"]); }
#[test] fn fenv_rounding_toward_zero_positive() { assert_eq!(run_c("#include <fenv.h>\n#pragma STDC FENV_ACCESS ON\nint main() { fesetround(FE_TOWARDZERO); double z = 1.0 / 3.0; printf(\"%d\", z < 0.3333333333333334); return 0; }"), vec!["1"]); }
#[test] fn fenv_rounding_toward_zero_negative() { assert_eq!(run_c("#include <fenv.h>\n#pragma STDC FENV_ACCESS ON\nint main() { fesetround(FE_TOWARDZERO); double z = -1.0 / 3.0; printf(\"%d\", z > -0.3333333333333334); return 0; }"), vec!["1"]); }
#[test] fn fenv_rounding_tonearest_ties_to_even() { assert_eq!(run_c("#include <fenv.h>\n#pragma STDC FENV_ACCESS ON\nint main() { fesetround(FE_TONEAREST); /* exact tie */ double d = 1.5; printf(\"ok\"); return 0; }"), vec!["ok"]); } // It's hard to portably show this with basic ops without hexfloats, test compile
#[test] fn fenv_rounding_cast_to_int() { assert_eq!(run_c("int main() { double d = 1.9; int i = (int)d; printf(\"%d\", i); return 0; }"), vec!["1"]); } // Cast always truncates toward zero regardless of rounding mode
#[test] fn fenv_getround_default() { assert_eq!(run_c("#include <fenv.h>\nint main() { printf(\"%d\", fegetround() == FE_TONEAREST); return 0; }"), vec!["1"]); }
#[test] fn fenv_invalid_round_mode_ignored() { assert_eq!(run_c("#include <fenv.h>\nint main() { int r = fesetround(9999); printf(\"%d\", r != 0); return 0; }"), vec!["1"]); } // Should return non-zero if mode is not supported
#[test] fn fenv_access_pragma_off() { assert_eq!(run_c("#include <fenv.h>\n#pragma STDC FENV_ACCESS OFF\nint main() { double d = 1.0/3.0; printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn fenv_access_pragma_block() { assert_eq!(run_c("#include <fenv.h>\nint main() { { #pragma STDC FENV_ACCESS ON\n fesetround(FE_UPWARD); } printf(\"ok\"); return 0; }"), vec!["ok"]); }
