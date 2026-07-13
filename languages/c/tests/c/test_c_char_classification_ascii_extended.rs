use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn ctype_isalpha_true() { assert_eq!(run_c("#include <ctype.h>\nint main() { printf(\"%d\", isalpha('A') != 0); return 0; }"), vec!["1"]); }
#[test] fn ctype_isalpha_false() { assert_eq!(run_c("#include <ctype.h>\nint main() { printf(\"%d\", isalpha('1') == 0); return 0; }"), vec!["1"]); }
#[test] fn ctype_isdigit_true() { assert_eq!(run_c("#include <ctype.h>\nint main() { printf(\"%d\", isdigit('5') != 0); return 0; }"), vec!["1"]); }
#[test] fn ctype_isalnum_true() { assert_eq!(run_c("#include <ctype.h>\nint main() { printf(\"%d\", isalnum('b') != 0 && isalnum('8') != 0); return 0; }"), vec!["1"]); }
#[test] fn ctype_isspace_space() { assert_eq!(run_c("#include <ctype.h>\nint main() { printf(\"%d\", isspace(' ') != 0); return 0; }"), vec!["1"]); }
#[test] fn ctype_isspace_newline() { assert_eq!(run_c("#include <ctype.h>\nint main() { printf(\"%d\", isspace('\\n') != 0); return 0; }"), vec!["1"]); }
#[test] fn ctype_isspace_tab() { assert_eq!(run_c("#include <ctype.h>\nint main() { printf(\"%d\", isspace('\\t') != 0); return 0; }"), vec!["1"]); }
#[test] fn ctype_islower_true() { assert_eq!(run_c("#include <ctype.h>\nint main() { printf(\"%d\", islower('a') != 0); return 0; }"), vec!["1"]); }
#[test] fn ctype_isupper_true() { assert_eq!(run_c("#include <ctype.h>\nint main() { printf(\"%d\", isupper('Z') != 0); return 0; }"), vec!["1"]); }
#[test] fn ctype_isxdigit_true() { assert_eq!(run_c("#include <ctype.h>\nint main() { printf(\"%d\", isxdigit('F') != 0 && isxdigit('a') != 0 && isxdigit('3') != 0); return 0; }"), vec!["1"]); }
#[test] fn ctype_iscntrl_true() { assert_eq!(run_c("#include <ctype.h>\nint main() { printf(\"%d\", iscntrl('\\n') != 0 && iscntrl(127) != 0); return 0; }"), vec!["1"]); }
#[test] fn ctype_isprint_true() { assert_eq!(run_c("#include <ctype.h>\nint main() { printf(\"%d\", isprint(' ') != 0 && isprint('A') != 0); return 0; }"), vec!["1"]); }
#[test] fn ctype_isgraph_true() { assert_eq!(run_c("#include <ctype.h>\nint main() { printf(\"%d\", isgraph(' ') == 0 && isgraph('A') != 0); return 0; }"), vec!["1"]); }
#[test] fn ctype_ispunct_true() { assert_eq!(run_c("#include <ctype.h>\nint main() { printf(\"%d\", ispunct(',') != 0 && ispunct('a') == 0); return 0; }"), vec!["1"]); }
#[test] fn ctype_tolower_basic() { assert_eq!(run_c("#include <ctype.h>\nint main() { printf(\"%c\", tolower('A')); return 0; }"), vec!["a"]); }
#[test] fn ctype_toupper_basic() { assert_eq!(run_c("#include <ctype.h>\nint main() { printf(\"%c\", toupper('a')); return 0; }"), vec!["A"]); }
#[test] fn ctype_tolower_nonalpha() { assert_eq!(run_c("#include <ctype.h>\nint main() { printf(\"%c\", tolower('1')); return 0; }"), vec!["1"]); }
