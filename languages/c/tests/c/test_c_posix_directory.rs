use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn posix_directory_lifecycle() {
    assert_eq!(
        run_c(
            r#"
#define _POSIX_C_SOURCE 200809L
#include <sys/stat.h>
#include <dirent.h>
#include <unistd.h>
#include <string.h>
#include <fcntl.h>

int main() {
    const char *dir_name = "test_dir_lifecycle";
    mkdir(dir_name, 0755);
    
    // Create a file inside
    char filepath[256];
    snprintf(filepath, sizeof(filepath), "%s/file1.txt", dir_name);
    int fd = open(filepath, O_CREAT | O_WRONLY, 0644);
    close(fd);
    
    DIR *d = opendir(dir_name);
    if (!d) return 1;
    
    int found_file = 0;
    struct dirent *dir;
    while ((dir = readdir(d)) != NULL) {
        if (strcmp(dir->d_name, "file1.txt") == 0) {
            found_file = 1;
        }
    }
    closedir(d);
    
    unlink(filepath);
    rmdir(dir_name);
    
    printf("%d", found_file);
    return 0;
}
    "#
        ),
        vec!["1"]
    );
}

#[test]
fn posix_getcwd_chdir() {
    assert_eq!(
        run_c(
            r#"
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
#include <stdlib.h>
#include <string.h>

int main() {
    char cwd1[1024];
    getcwd(cwd1, sizeof(cwd1));
    
    chdir("/");
    char cwd2[1024];
    getcwd(cwd2, sizeof(cwd2));
    
    chdir(cwd1);
    char cwd3[1024];
    getcwd(cwd3, sizeof(cwd3));
    
    printf("%d %d", strcmp(cwd2, "/") == 0, strcmp(cwd1, cwd3) == 0);
    return 0;
}
    "#
        ),
        vec!["1 1"]
    );
}
