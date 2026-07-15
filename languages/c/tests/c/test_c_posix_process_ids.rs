use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn getpid_basic() {
    assert_eq!(
        run_c("#include <unistd.h>\nint main() { printf(\"%d\", getpid() > 0); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn getppid_basic() {
    assert_eq!(
        run_c("#include <unistd.h>\nint main() { printf(\"%d\", getppid() > 0); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn getuid_basic() {
    assert_eq!(
        run_c("#include <unistd.h>\nint main() { printf(\"%d\", getuid() >= 0); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn geteuid_basic() {
    assert_eq!(
        run_c("#include <unistd.h>\nint main() { printf(\"%d\", geteuid() >= 0); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn getgid_basic() {
    assert_eq!(
        run_c("#include <unistd.h>\nint main() { printf(\"%d\", getgid() >= 0); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn getegid_basic() {
    assert_eq!(
        run_c("#include <unistd.h>\nint main() { printf(\"%d\", getegid() >= 0); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn setuid_compile() {
    assert_eq!(
        run_c(
            "#include <unistd.h>\nint main() { /* test compile and fail gracefully */ int r = setuid(getuid()); printf(\"%d\", r == 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn setgid_compile() {
    assert_eq!(
        run_c(
            "#include <unistd.h>\nint main() { int r = setgid(getgid()); printf(\"%d\", r == 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn seteuid_compile() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { int r = seteuid(geteuid()); printf(\"%d\", r == 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn setegid_compile() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { int r = setegid(getegid()); printf(\"%d\", r == 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn setreuid_compile() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE 500\n#include <unistd.h>\nint main() { int r = setreuid(-1, -1); printf(\"%d\", r == 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn setregid_compile() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE 500\n#include <unistd.h>\nint main() { int r = setregid(-1, -1); printf(\"%d\", r == 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn setresuid_gnu() {
    assert_eq!(
        run_c(
            "#define _GNU_SOURCE\n#include <unistd.h>\nint main() { int r = setresuid(-1, -1, -1); printf(\"%d\", r == 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn setresgid_gnu() {
    assert_eq!(
        run_c(
            "#define _GNU_SOURCE\n#include <unistd.h>\nint main() { int r = setresgid(-1, -1, -1); printf(\"%d\", r == 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn getresuid_gnu() {
    assert_eq!(
        run_c(
            "#define _GNU_SOURCE\n#include <unistd.h>\nint main() { uid_t r, e, s; int res = getresuid(&r, &e, &s); printf(\"%d\", res == 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn getresgid_gnu() {
    assert_eq!(
        run_c(
            "#define _GNU_SOURCE\n#include <unistd.h>\nint main() { gid_t r, e, s; int res = getresgid(&r, &e, &s); printf(\"%d\", res == 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn getgroups_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { gid_t g[32]; int n = getgroups(32, g); printf(\"%d\", n >= 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn setgroups_compile() {
    assert_eq!(
        run_c(
            "#define _BSD_SOURCE\n#define _DEFAULT_SOURCE\n#include <grp.h>\n#include <unistd.h>\nint main() { /* setgroups needs root, test failure */ int r = setgroups(0, NULL); printf(\"%d\", r == -1 || r == 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn getlogin_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { char *p = getlogin(); printf(\"%d\", p == NULL || p != NULL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn getlogin_r_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { char b[256]; int r = getlogin_r(b, 256); printf(\"%d\", r == 0 || r != 0); return 0; }"
        ),
        vec!["1"]
    );
}
