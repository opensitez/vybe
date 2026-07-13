use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn posix_stat_complex_workflow() {
    assert_eq!(run_c(r#"
#define _POSIX_C_SOURCE 200809L
#include <sys/stat.h>
#include <sys/types.h>
#include <unistd.h>
#include <fcntl.h>
#include <stdlib.h>

int main() {
    const char *path = "test_stat_complex.txt";
    int fd = open(path, O_CREAT | O_WRONLY | O_TRUNC, 0644);
    if (fd < 0) return 1;
    write(fd, "hello world", 11);
    close(fd);

    struct stat st;
    if (stat(path, &st) == 0) {
        printf("%d %d %d", (int)st.st_size, S_ISREG(st.st_mode), S_ISDIR(st.st_mode));
    }
    unlink(path);
    return 0;
}
    "#), vec!["11 1 0"]);
}

#[test] fn posix_fstat_lstat_comparison() {
    assert_eq!(run_c(r#"
#define _POSIX_C_SOURCE 200809L
#include <sys/stat.h>
#include <unistd.h>
#include <fcntl.h>

int main() {
    const char *path = "test_fstat_lstat.txt";
    int fd = open(path, O_CREAT | O_WRONLY, 0644);
    if (fd < 0) return 1;
    write(fd, "X", 1);
    
    struct stat st1, st2;
    fstat(fd, &st1);
    lstat(path, &st2);
    close(fd);
    unlink(path);
    
    printf("%d %d", (int)st1.st_size == (int)st2.st_size, st1.st_mode == st2.st_mode);
    return 0;
}
    "#), vec!["1 1"]);
}

#[test] fn posix_access_permissions() {
    assert_eq!(run_c(r#"
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
#include <fcntl.h>

int main() {
    const char *path = "test_access.txt";
    int fd = open(path, O_CREAT | O_WRONLY, 0600);
    close(fd);
    
    int r_ok = access(path, R_OK) == 0;
    int w_ok = access(path, W_OK) == 0;
    int x_ok = access(path, X_OK) == 0;
    
    unlink(path);
    printf("%d %d %d", r_ok, w_ok, x_ok);
    return 0;
}
    "#), vec!["1 1 0"]);
}
