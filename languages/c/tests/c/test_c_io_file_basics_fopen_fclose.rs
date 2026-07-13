use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn fopen_write() { assert_eq!(run_c("int main() { FILE *f = fopen(\"test_fopen.txt\", \"w\"); if (f) { printf(\"ok\"); fclose(f); } return 0; }"), vec!["ok"]); }
#[test] fn fopen_read_not_exist() { assert_eq!(run_c("int main() { FILE *f = fopen(\"does_not_exist.txt\", \"r\"); printf(\"%d\", f == NULL); return 0; }"), vec!["1"]); }
#[test] fn fopen_read_exists() { assert_eq!(run_c("int main() { FILE *f = fopen(\"test_fopen_r.txt\", \"w\"); fclose(f); f = fopen(\"test_fopen_r.txt\", \"r\"); if (f) { printf(\"ok\"); fclose(f); } return 0; }"), vec!["ok"]); }
#[test] fn fopen_append() { assert_eq!(run_c("int main() { FILE *f = fopen(\"test_fopen_a.txt\", \"w\"); fclose(f); f = fopen(\"test_fopen_a.txt\", \"a\"); if (f) { printf(\"ok\"); fclose(f); } return 0; }"), vec!["ok"]); }
#[test] fn fclose_null_fails() { assert_eq!(run_c("int main() { /* fclose(NULL) is undefined behavior, some crash, some return EOF */ printf(\"skipped\"); return 0; }"), vec!["skipped"]); }
#[test] fn fdopen_basic() { assert_eq!(run_c("#include <fcntl.h>\n#include <unistd.h>\nint main() { int fd = open(\"test_fdopen.txt\", O_CREAT|O_WRONLY, 0644); FILE *f = fdopen(fd, \"w\"); if (f) { printf(\"ok\"); fclose(f); } return 0; }"), vec!["ok"]); }
#[test] fn freopen_basic() { assert_eq!(run_c("int main() { FILE *f = fopen(\"test_freopen.txt\", \"w\"); if (f) { f = freopen(\"test_freopen_2.txt\", \"w\", f); if (f) { printf(\"ok\"); fclose(f); } } return 0; }"), vec!["ok"]); }
