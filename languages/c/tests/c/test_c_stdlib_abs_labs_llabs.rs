use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn abs_positive() { assert_eq!(run_c("#include <stdlib.h>\nint main() { printf(\"%d\", abs(42)); return 0; }"), vec!["42"]); }
#[test] fn abs_negative() { assert_eq!(run_c("#include <stdlib.h>\nint main() { printf(\"%d\", abs(-42)); return 0; }"), vec!["42"]); }
#[test] fn abs_zero() { assert_eq!(run_c("#include <stdlib.h>\nint main() { printf(\"%d\", abs(0)); return 0; }"), vec!["0"]); }
#[test] fn labs_positive() { assert_eq!(run_c("#include <stdlib.h>\nint main() { printf(\"%ld\", labs(42L)); return 0; }"), vec!["42"]); }
#[test] fn labs_negative() { assert_eq!(run_c("#include <stdlib.h>\nint main() { printf(\"%ld\", labs(-42L)); return 0; }"), vec!["42"]); }
#[test] fn labs_zero() { assert_eq!(run_c("#include <stdlib.h>\nint main() { printf(\"%ld\", labs(0L)); return 0; }"), vec!["0"]); }
#[test] fn llabs_positive() { assert_eq!(run_c("#include <stdlib.h>\nint main() { printf(\"%lld\", llabs(42LL)); return 0; }"), vec!["42"]); }
#[test] fn llabs_negative() { assert_eq!(run_c("#include <stdlib.h>\nint main() { printf(\"%lld\", llabs(-42LL)); return 0; }"), vec!["42"]); }
#[test] fn llabs_zero() { assert_eq!(run_c("#include <stdlib.h>\nint main() { printf(\"%lld\", llabs(0LL)); return 0; }"), vec!["0"]); }
#[test] fn imaxabs_positive() { assert_eq!(run_c("#include <inttypes.h>\nint main() { printf(\"%jd\", imaxabs((intmax_t)42)); return 0; }"), vec!["42"]); }
#[test] fn imaxabs_negative() { assert_eq!(run_c("#include <inttypes.h>\nint main() { printf(\"%jd\", imaxabs((intmax_t)-42)); return 0; }"), vec!["42"]); }
#[test] fn imaxabs_zero() { assert_eq!(run_c("#include <inttypes.h>\nint main() { printf(\"%jd\", imaxabs((intmax_t)0)); return 0; }"), vec!["0"]); }
#[test] fn abs_limits() { assert_eq!(run_c("#include <stdlib.h>\n#include <limits.h>\nint main() { printf(\"%d\", abs(-INT_MAX)); return 0; }"), vec![format!("{}", i32::MAX)]); }
#[test] fn labs_limits() { assert_eq!(run_c("#include <stdlib.h>\n#include <limits.h>\nint main() { printf(\"%ld\", labs(-LONG_MAX)); return 0; }"), vec![format!("{}", i64::MAX)]); } // Assuming 64-bit long
#[test] fn llabs_limits() { assert_eq!(run_c("#include <stdlib.h>\n#include <limits.h>\nint main() { printf(\"%lld\", llabs(-LLONG_MAX)); return 0; }"), vec![format!("{}", i64::MAX)]); }
#[test] fn imaxabs_limits() { assert_eq!(run_c("#include <inttypes.h>\n#include <limits.h>\nint main() { printf(\"%jd\", imaxabs(-INTMAX_MAX)); return 0; }"), vec![format!("{}", i64::MAX)]); } // Assuming intmax_t is 64-bit
#[test] fn abs_expression() { assert_eq!(run_c("#include <stdlib.h>\nint main() { int x = 10, y = 20; printf(\"%d\", abs(x - y)); return 0; }"), vec!["10"]); }
#[test] fn labs_expression() { assert_eq!(run_c("#include <stdlib.h>\nint main() { long x = 10L, y = 20L; printf(\"%ld\", labs(x - y)); return 0; }"), vec!["10"]); }
#[test] fn llabs_expression() { assert_eq!(run_c("#include <stdlib.h>\nint main() { long long x = 10LL, y = 20LL; printf(\"%lld\", llabs(x - y)); return 0; }"), vec!["10"]); }
#[test] fn imaxabs_expression() { assert_eq!(run_c("#include <inttypes.h>\nint main() { intmax_t x = 10, y = 20; printf(\"%jd\", imaxabs(x - y)); return 0; }"), vec!["10"]); }
