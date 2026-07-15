use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn div_basic() { assert_eq!(run_c("#include <stdlib.h>\nint main() { div_t d = div(10, 3); printf(\"%d %d\", d.quot, d.rem); return 0; }"), vec!["3 1"]); }
#[test] fn div_negative_numerator() { assert_eq!(run_c("#include <stdlib.h>\nint main() { div_t d = div(-10, 3); printf(\"%d %d\", d.quot, d.rem); return 0; }"), vec!["-3 -1"]); } // C99 mandates truncation towards zero
#[test] fn div_negative_denominator() { assert_eq!(run_c("#include <stdlib.h>\nint main() { div_t d = div(10, -3); printf(\"%d %d\", d.quot, d.rem); return 0; }"), vec!["-3 1"]); }
#[test] fn div_both_negative() { assert_eq!(run_c("#include <stdlib.h>\nint main() { div_t d = div(-10, -3); printf(\"%d %d\", d.quot, d.rem); return 0; }"), vec!["3 -1"]); }
#[test] fn div_exact() { assert_eq!(run_c("#include <stdlib.h>\nint main() { div_t d = div(9, 3); printf(\"%d %d\", d.quot, d.rem); return 0; }"), vec!["3 0"]); }
#[test] fn ldiv_basic() { assert_eq!(run_c("#include <stdlib.h>\nint main() { ldiv_t d = ldiv(100L, 30L); printf(\"%ld %ld\", d.quot, d.rem); return 0; }"), vec!["3 10"]); }
#[test] fn ldiv_negative() { assert_eq!(run_c("#include <stdlib.h>\nint main() { ldiv_t d = ldiv(-100L, 30L); printf(\"%ld %ld\", d.quot, d.rem); return 0; }"), vec!["-3 -10"]); }
#[test] fn lldiv_basic() { assert_eq!(run_c("#include <stdlib.h>\nint main() { lldiv_t d = lldiv(1000LL, 300LL); printf(\"%lld %lld\", d.quot, d.rem); return 0; }"), vec!["3 100"]); }
#[test] fn lldiv_negative() { assert_eq!(run_c("#include <stdlib.h>\nint main() { lldiv_t d = lldiv(1000LL, -300LL); printf(\"%lld %lld\", d.quot, d.rem); return 0; }"), vec!["-3 100"]); }
#[test] fn imaxdiv_basic() { assert_eq!(run_c("#include <inttypes.h>\nint main() { imaxdiv_t d = imaxdiv((intmax_t)10, (intmax_t)3); printf(\"%d %d\", (int)d.quot, (int)d.rem); return 0; }"), vec!["3 1"]); }
#[test] fn div_zero_numerator() { assert_eq!(run_c("#include <stdlib.h>\nint main() { div_t d = div(0, 5); printf(\"%d %d\", d.quot, d.rem); return 0; }"), vec!["0 0"]); }
#[test] fn ldiv_zero_numerator() { assert_eq!(run_c("#include <stdlib.h>\nint main() { ldiv_t d = ldiv(0L, 5L); printf(\"%ld %ld\", d.quot, d.rem); return 0; }"), vec!["0 0"]); }
#[test] fn lldiv_zero_numerator() { assert_eq!(run_c("#include <stdlib.h>\nint main() { lldiv_t d = lldiv(0LL, 5LL); printf(\"%lld %lld\", d.quot, d.rem); return 0; }"), vec!["0 0"]); }
#[test] fn imaxdiv_zero_numerator() { assert_eq!(run_c("#include <inttypes.h>\nint main() { imaxdiv_t d = imaxdiv(0, 5); printf(\"%d %d\", (int)d.quot, (int)d.rem); return 0; }"), vec!["0 0"]); }
#[test] fn div_limits() { assert_eq!(run_c("#include <stdlib.h>\n#include <limits.h>\nint main() { div_t d = div(INT_MAX, 2); printf(\"%d %d\", d.quot, d.rem); return 0; }"), vec![format!("{} 1", i32::MAX / 2)]); }
#[test] fn ldiv_limits() { assert_eq!(run_c("#include <stdlib.h>\n#include <limits.h>\nint main() { ldiv_t d = ldiv(LONG_MAX, 2L); printf(\"%ld %ld\", d.quot, d.rem); return 0; }"), vec![format!("{} 1", i64::MAX / 2)]); } // Assume 64 bit long
#[test] fn lldiv_limits() { assert_eq!(run_c("#include <stdlib.h>\n#include <limits.h>\nint main() { lldiv_t d = lldiv(LLONG_MAX, 2LL); printf(\"%lld %lld\", d.quot, d.rem); return 0; }"), vec![format!("{} 1", i64::MAX / 2)]); }
#[test] fn div_by_one() { assert_eq!(run_c("#include <stdlib.h>\nint main() { div_t d = div(42, 1); printf(\"%d %d\", d.quot, d.rem); return 0; }"), vec!["42 0"]); }
#[test] fn ldiv_by_one() { assert_eq!(run_c("#include <stdlib.h>\nint main() { ldiv_t d = ldiv(42L, 1L); printf(\"%ld %ld\", d.quot, d.rem); return 0; }"), vec!["42 0"]); }
#[test] fn lldiv_by_one() { assert_eq!(run_c("#include <stdlib.h>\nint main() { lldiv_t d = lldiv(42LL, 1LL); printf(\"%lld %lld\", d.quot, d.rem); return 0; }"), vec!["42 0"]); }
