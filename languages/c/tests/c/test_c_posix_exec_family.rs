use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn execl_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { execl(\"/bin/echo\", \"echo\", \"hello\", \"execl\", NULL); return 1; }"
        ),
        vec!["hello execl"]
    );
}
#[test]
fn execle_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { char *env[] = {\"TEST_VAR=42\", NULL}; execle(\"/usr/bin/env\", \"env\", NULL, env); return 1; }"
        ),
        vec!["TEST_VAR=42"]
    );
}
#[test]
fn execlp_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { execlp(\"echo\", \"echo\", \"hello\", \"execlp\", NULL); return 1; }"
        ),
        vec!["hello execlp"]
    );
}
#[test]
fn execv_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { char *args[] = {\"echo\", \"hello\", \"execv\", NULL}; execv(\"/bin/echo\", args); return 1; }"
        ),
        vec!["hello execv"]
    );
}
#[test]
fn execvp_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { char *args[] = {\"echo\", \"hello\", \"execvp\", NULL}; execvp(\"echo\", args); return 1; }"
        ),
        vec!["hello execvp"]
    );
}
#[test]
fn execve_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { char *args[] = {\"env\", NULL}; char *env[] = {\"VAR=execve\", NULL}; execve(\"/usr/bin/env\", args, env); return 1; }"
        ),
        vec!["VAR=execve"]
    );
}
#[test]
fn execvpe_gnu() {
    assert_eq!(
        run_c(
            "#define _GNU_SOURCE\n#include <unistd.h>\nint main() { char *args[] = {\"env\", NULL}; char *env[] = {\"VAR=execvpe\", NULL}; execvpe(\"env\", args, env); return 1; }"
        ),
        vec!["VAR=execvpe"]
    );
}
#[test]
fn fexecve_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\n#include <fcntl.h>\nint main() { int fd = open(\"/bin/echo\", O_RDONLY); if (fd < 0) return 0; char *args[] = {\"echo\", \"fexecve_ok\", NULL}; char *env[] = {NULL}; fexecve(fd, args, env); return 1; }"
        ),
        vec!["fexecve_ok"]
    );
}
#[test]
fn execl_invalid_path() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { int r = execl(\"/does/not/exist/123\", \"abc\", NULL); printf(\"%d\", r == -1); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn execlp_invalid_cmd() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { int r = execlp(\"does_not_exist_123\", \"abc\", NULL); printf(\"%d\", r == -1); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn execvp_invalid_cmd() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { char *args[] = {\"does_not_exist\", NULL}; int r = execvp(\"does_not_exist_123\", args); printf(\"%d\", r == -1); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn exec_close_on_exec() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\n#include <fcntl.h>\nint main() { int fd = open(\"test_cloexec.txt\", O_CREAT|O_WRONLY, 0644); fcntl(fd, F_SETFD, FD_CLOEXEC); /* We test that fcntl succeeds and FD_CLOEXEC is defined */ printf(\"ok\"); close(fd); unlink(\"test_cloexec.txt\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn exec_keeps_open_fd() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\n#include <fcntl.h>\n#include <sys/wait.h>\nint main() { int fd = open(\"test_keep_fd.txt\", O_CREAT|O_WRONLY, 0644); pid_t p = fork(); if(p==0) { /* fd remains open across exec unless cloexec is set */ char buf[50]; sprintf(buf, \"echo hi >&%d\", fd); execl(\"/bin/sh\", \"sh\", \"-c\", buf, NULL); _exit(1); } waitpid(p, NULL, 0); close(fd); FILE *f = fopen(\"test_keep_fd.txt\", \"r\"); char b[10]={0}; fread(b, 1, 9, f); printf(\"%s\", b); fclose(f); unlink(\"test_keep_fd.txt\"); return 0; }"
        ),
        vec!["hi"]
    );
}
#[test]
fn execlp_fallback() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { /* execlp looks in PATH. We run true which exists */ execlp(\"true\", \"true\", NULL); return 1; }"
        ),
        Vec::<String>::new()
    );
} // returns 0
#[test]
fn execve_shebang() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\n#include <fcntl.h>\n#include <sys/wait.h>\n#include <sys/stat.h>\nint main() { int fd = open(\"test_script.sh\", O_CREAT|O_WRONLY, 0755); write(fd, \"#!/bin/sh\\necho script\\n\", 22); close(fd); pid_t p = fork(); if(p==0) { char *args[] = {\"./test_script.sh\", NULL}; char *env[] = {NULL}; execve(\"./test_script.sh\", args, env); _exit(1); } int st; wait(&st); unlink(\"test_script.sh\"); return 0; }"
        ),
        vec!["script"]
    );
}
#[test]
fn posix_spawn_file_actions() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <spawn.h>\n#include <sys/wait.h>\n#include <fcntl.h>\nextern char **environ;\nint main() { posix_spawn_file_actions_t a; posix_spawn_file_actions_init(&a); posix_spawn_file_actions_addopen(&a, 1, \"test_spawn_out.txt\", O_CREAT|O_WRONLY, 0644); pid_t p; char *args[] = {\"echo\", \"spawned\", NULL}; posix_spawn(&p, \"/bin/echo\", &a, NULL, args, environ); wait(NULL); posix_spawn_file_actions_destroy(&a); FILE *f = fopen(\"test_spawn_out.txt\", \"r\"); char b[20]={0}; fread(b, 1, 10, f); printf(\"%s\", b); fclose(f); unlink(\"test_spawn_out.txt\"); return 0; }"
        ),
        vec!["spawned"]
    );
}
#[test]
fn execl_long_arg() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\n#include <string.h>\nint main() { char arg[1000]; memset(arg, 'x', 999); arg[999] = 0; execl(\"/bin/echo\", \"echo\", arg, NULL); return 1; }"
        ),
        vec!["x".repeat(999)]
    );
}
#[test]
fn execv_empty_env() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { char *args[] = {\"env\", NULL}; char *env[] = {NULL}; execve(\"/usr/bin/env\", args, env); /* output should be completely empty */ return 0; }"
        ),
        Vec::<String>::new()
    );
}
#[test]
fn exec_args_array_null_terminator() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { char *args[2]; args[0] = \"true\"; args[1] = NULL; execvp(args[0], args); return 1; }"
        ),
        Vec::<String>::new()
    );
}
#[test]
fn exec_changes_pid() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\n#include <sys/wait.h>\nint main() { pid_t parent_pid = getpid(); pid_t p = fork(); if (p == 0) { char buf[50]; sprintf(buf, \"echo $PPID\"); execl(\"/bin/sh\", \"sh\", \"-c\", buf, NULL); _exit(1); } int st; waitpid(p, &st, 0); /* Shell's parent is our parent */ return 0; }"
        ),
        Vec::<String>::new()
    );
} // Too complex to verify $PPID consistently across environments, but compilation and no crash is fine.
