use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn strchr_found() { assert_eq!(run_c("#include <string.h>\nint main() { char *s = \"hello\"; printf(\"%s\", strchr(s, 'e')); return 0; }"), vec!["ello"]); }
#[test] fn strchr_not_found() { assert_eq!(run_c("#include <string.h>\nint main() { char *s = \"hello\"; printf(\"%d\", strchr(s, 'x') == NULL); return 0; }"), vec!["1"]); }
#[test] fn strchr_null_char() { assert_eq!(run_c("#include <string.h>\nint main() { char *s = \"hello\"; printf(\"%d\", strchr(s, '\\0') == s + 5); return 0; }"), vec!["1"]); }
#[test] fn strrchr_found() { assert_eq!(run_c("#include <string.h>\nint main() { char *s = \"hello\"; printf(\"%s\", strrchr(s, 'l')); return 0; }"), vec!["lo"]); }
#[test] fn strrchr_not_found() { assert_eq!(run_c("#include <string.h>\nint main() { char *s = \"hello\"; printf(\"%d\", strrchr(s, 'x') == NULL); return 0; }"), vec!["1"]); }
#[test] fn strrchr_null_char() { assert_eq!(run_c("#include <string.h>\nint main() { char *s = \"hello\"; printf(\"%d\", strrchr(s, '\\0') == s + 5); return 0; }"), vec!["1"]); }
#[test] fn strstr_found() { assert_eq!(run_c("#include <string.h>\nint main() { char *s = \"hello world\"; printf(\"%s\", strstr(s, \"lo\")); return 0; }"), vec!["lo world"]); }
#[test] fn strstr_not_found() { assert_eq!(run_c("#include <string.h>\nint main() { char *s = \"hello world\"; printf(\"%d\", strstr(s, \"xyz\") == NULL); return 0; }"), vec!["1"]); }
#[test] fn strstr_empty_needle() { assert_eq!(run_c("#include <string.h>\nint main() { char *s = \"hello\"; printf(\"%s\", strstr(s, \"\")); return 0; }"), vec!["hello"]); }
#[test] fn strstr_needle_longer_than_haystack() { assert_eq!(run_c("#include <string.h>\nint main() { char *s = \"hi\"; printf(\"%d\", strstr(s, \"hello\") == NULL); return 0; }"), vec!["1"]); }
#[test] fn strcasestr_found_gnu() { assert_eq!(run_c("#define _GNU_SOURCE\n#include <string.h>\nint main() { char *s = \"Hello World\"; printf(\"%s\", strcasestr(s, \"LO\")); return 0; }"), vec!["lo World"]); }
#[test] fn strcasestr_not_found_gnu() { assert_eq!(run_c("#define _GNU_SOURCE\n#include <string.h>\nint main() { char *s = \"Hello World\"; printf(\"%d\", strcasestr(s, \"xyz\") == NULL); return 0; }"), vec!["1"]); }
#[test] fn memchr_found() { assert_eq!(run_c("#include <string.h>\nint main() { char s[] = {1, 2, 3, 4}; printf(\"%d\", (int)((char*)memchr(s, 3, 4) - s)); return 0; }"), vec!["2"]); }
#[test] fn memchr_not_found() { assert_eq!(run_c("#include <string.h>\nint main() { char s[] = {1, 2, 3, 4}; printf(\"%d\", memchr(s, 5, 4) == NULL); return 0; }"), vec!["1"]); }
#[test] fn memrchr_found_gnu() { assert_eq!(run_c("#define _GNU_SOURCE\n#include <string.h>\nint main() { char s[] = {1, 2, 1, 4}; printf(\"%d\", (int)((char*)memrchr(s, 1, 4) - s)); return 0; }"), vec!["2"]); }
