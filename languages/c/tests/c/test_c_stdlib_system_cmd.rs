use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn system_echo() { assert_eq!(run_c("#include <stdlib.h>\nint main() { system(\"echo hello\"); return 0; }"), vec!["hello"]); }
#[test] fn system_null() { assert_eq!(run_c("#include <stdlib.h>\nint main() { int res = system(NULL); printf(\"%d\", res != 0); return 0; }"), vec!["1"]); } // Returns non-zero if shell is available
#[test] fn system_return_code() { assert_eq!(run_c("#include <stdlib.h>\nint main() { int res = system(\"exit 42\"); printf(\"%d\", WEXITSTATUS(res)); return 0; }"), vec!["42"]); }
#[test] fn system_invalid_cmd() { assert_eq!(run_c("#include <stdlib.h>\nint main() { int res = system(\"nonexistent_command_12345 2>/dev/null\"); printf(\"%d\", res != 0); return 0; }"), vec!["1"]); }
#[test] fn system_background() { assert_eq!(run_c("#include <stdlib.h>\nint main() { int res = system(\"sleep 0.1 &\"); printf(\"%d\", res == 0); return 0; }"), vec!["1"]); }
#[test] fn system_pipe() { assert_eq!(run_c("#include <stdlib.h>\nint main() { system(\"echo abc | tr a-z A-Z\"); return 0; }"), vec!["ABC"]); }
#[test] fn system_redirection_stdout() { assert_eq!(run_c("#include <stdlib.h>\nint main() { system(\"echo abc > test_system_redirect.txt\"); FILE *f = fopen(\"test_system_redirect.txt\", \"r\"); char buf[10]; fgets(buf, sizeof(buf), f); printf(\"%s\", buf); fclose(f); system(\"rm test_system_redirect.txt\"); return 0; }"), vec!["abc\n"]); }
#[test] fn system_redirection_stderr() { assert_eq!(run_c("#include <stdlib.h>\nint main() { system(\"ls nonexistent_file 2> test_system_err.txt\"); FILE *f = fopen(\"test_system_err.txt\", \"r\"); char buf[10]; printf(\"%d\", fgets(buf, sizeof(buf), f) != NULL); fclose(f); system(\"rm test_system_err.txt\"); return 0; }"), vec!["1"]); }
#[test] fn system_multiple_commands() { assert_eq!(run_c("#include <stdlib.h>\nint main() { system(\"echo 1; echo 2; echo 3\"); return 0; }"), vec!["1\n2\n3"]); }
#[test] fn system_logical_and() { assert_eq!(run_c("#include <stdlib.h>\nint main() { system(\"true && echo ok\"); return 0; }"), vec!["ok"]); }
#[test] fn system_logical_or() { assert_eq!(run_c("#include <stdlib.h>\nint main() { system(\"false || echo ok\"); return 0; }"), vec!["ok"]); }
#[test] fn system_subshell() { assert_eq!(run_c("#include <stdlib.h>\nint main() { system(\"(echo sub)\"); return 0; }"), vec!["sub"]); }
#[test] fn system_environment() { assert_eq!(run_c("#include <stdlib.h>\nint main() { setenv(\"SYS_VAR\", \"123\", 1); system(\"echo $SYS_VAR\"); return 0; }"), vec!["123"]); }
#[test] fn system_empty_cmd() { assert_eq!(run_c("#include <stdlib.h>\nint main() { int res = system(\"\"); printf(\"%d\", res == 0); return 0; }"), vec!["1"]); }
#[test] fn system_signals() { assert_eq!(run_c("#include <stdlib.h>\n#include <signal.h>\nint main() { /* system() ignores SIGINT/SIGQUIT and blocks SIGCHLD. Hard to test output directly, but we can verify it returns properly */ int res = system(\"kill -INT $$\"); printf(\"%d\", res != 0); return 0; }"), vec!["1"]); }
#[test] fn system_wait_for_child() { assert_eq!(run_c("#include <stdlib.h>\nint main() { system(\"sleep 0.1\"); printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn system_command_substitution() { assert_eq!(run_c("#include <stdlib.h>\nint main() { system(\"echo $(echo nested)\"); return 0; }"), vec!["nested"]); }
#[test] fn system_wildcards() { assert_eq!(run_c("#include <stdlib.h>\nint main() { system(\"echo test_c_*.rs > /dev/null\"); printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn system_quotes() { assert_eq!(run_c("#include <stdlib.h>\nint main() { system(\"echo 'single quotes'\"); system(\"echo \\\"double quotes\\\"\"); return 0; }"), vec!["single quotes\ndouble quotes"]); }
#[test] fn system_escapes() { assert_eq!(run_c("#include <stdlib.h>\nint main() { system(\"echo \\\\$\\\\$\"); return 0; }"), vec!["$$"]); }
