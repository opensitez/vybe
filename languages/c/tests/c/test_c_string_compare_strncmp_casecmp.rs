use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn strncmp_equal() { assert_eq!(run_c("#include <string.h>\nint main() { printf(\"%d\", strncmp(\"hello\", \"hello\", 5) == 0); return 0; }"), vec!["1"]); }
#[test] fn strncmp_less() { assert_eq!(run_c("#include <string.h>\nint main() { printf(\"%d\", strncmp(\"hella\", \"hello\", 5) < 0); return 0; }"), vec!["1"]); }
#[test] fn strncmp_greater() { assert_eq!(run_c("#include <string.h>\nint main() { printf(\"%d\", strncmp(\"hellz\", \"hello\", 5) > 0); return 0; }"), vec!["1"]); }
#[test] fn strncmp_equal_prefix() { assert_eq!(run_c("#include <string.h>\nint main() { printf(\"%d\", strncmp(\"hello\", \"hellz\", 4) == 0); return 0; }"), vec!["1"]); }
#[test] fn strncmp_zero_length() { assert_eq!(run_c("#include <string.h>\nint main() { printf(\"%d\", strncmp(\"a\", \"b\", 0) == 0); return 0; }"), vec!["1"]); }
#[test] fn strncmp_null_termination_early() { assert_eq!(run_c("#include <string.h>\nint main() { printf(\"%d\", strncmp(\"hi\", \"hi\", 10) == 0); return 0; }"), vec!["1"]); }
#[test] fn strcasecmp_equal() { assert_eq!(run_c("#include <strings.h>\nint main() { printf(\"%d\", strcasecmp(\"HeLlO\", \"hElLo\") == 0); return 0; }"), vec!["1"]); }
#[test] fn strcasecmp_less() { assert_eq!(run_c("#include <strings.h>\nint main() { printf(\"%d\", strcasecmp(\"apple\", \"BANANA\") < 0); return 0; }"), vec!["1"]); }
#[test] fn strcasecmp_greater() { assert_eq!(run_c("#include <strings.h>\nint main() { printf(\"%d\", strcasecmp(\"Zebra\", \"apple\") > 0); return 0; }"), vec!["1"]); }
#[test] fn strncasecmp_equal_prefix() { assert_eq!(run_c("#include <strings.h>\nint main() { printf(\"%d\", strncasecmp(\"HeLlO\", \"hElLz\", 4) == 0); return 0; }"), vec!["1"]); }
#[test] fn strncasecmp_diff_after_prefix() { assert_eq!(run_c("#include <strings.h>\nint main() { printf(\"%d\", strncasecmp(\"HeLlO\", \"hElLz\", 5) < 0); return 0; }"), vec!["1"]); }
#[test] fn memcmp_equal() { assert_eq!(run_c("#include <string.h>\nint main() { char a[] = {1, 2, 3}; char b[] = {1, 2, 3}; printf(\"%d\", memcmp(a, b, 3) == 0); return 0; }"), vec!["1"]); }
#[test] fn memcmp_less() { assert_eq!(run_c("#include <string.h>\nint main() { char a[] = {1, 2, 3}; char b[] = {1, 3, 3}; printf(\"%d\", memcmp(a, b, 3) < 0); return 0; }"), vec!["1"]); }
#[test] fn memcmp_zero_length() { assert_eq!(run_c("#include <string.h>\nint main() { char a[] = {1}; char b[] = {2}; printf(\"%d\", memcmp(a, b, 0) == 0); return 0; }"), vec!["1"]); }
