use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn strlen_basic() { assert_eq!(run_c("#include <string.h>\nint main() { printf(\"%d\", (int)strlen(\"hello\")); return 0; }"), vec!["5"]); }
#[test] fn strlen_empty() { assert_eq!(run_c("#include <string.h>\nint main() { printf(\"%d\", (int)strlen(\"\")); return 0; }"), vec!["0"]); }
#[test] fn strnlen_basic() { assert_eq!(run_c("#include <string.h>\nint main() { printf(\"%d\", (int)strnlen(\"hello\", 10)); return 0; }"), vec!["5"]); }
#[test] fn strnlen_truncation() { assert_eq!(run_c("#include <string.h>\nint main() { printf(\"%d\", (int)strnlen(\"hello\", 3)); return 0; }"), vec!["3"]); }
#[test] fn strnlen_zero_max() { assert_eq!(run_c("#include <string.h>\nint main() { printf(\"%d\", (int)strnlen(\"hello\", 0)); return 0; }"), vec!["0"]); }
#[test] fn strnlen_no_null_term() { assert_eq!(run_c("#include <string.h>\nint main() { char s[3] = {'a', 'b', 'c'}; printf(\"%d\", (int)strnlen(s, 3)); return 0; }"), vec!["3"]); }
