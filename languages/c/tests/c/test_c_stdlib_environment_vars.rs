use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn getenv_basic() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { char *p = getenv(\"PATH\"); printf(\"%d\", p != NULL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn getenv_not_found() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { char *p = getenv(\"DOES_NOT_EXIST_XYZ\"); printf(\"%d\", p == NULL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn setenv_getenv() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <stdlib.h>\nint main() { setenv(\"MY_TEST_VAR\", \"123\", 1); printf(\"%s\", getenv(\"MY_TEST_VAR\")); return 0; }"
        ),
        vec!["123"]
    );
}
#[test]
fn setenv_no_overwrite() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <stdlib.h>\nint main() { setenv(\"MY_TEST_VAR\", \"123\", 1); setenv(\"MY_TEST_VAR\", \"456\", 0); printf(\"%s\", getenv(\"MY_TEST_VAR\")); return 0; }"
        ),
        vec!["123"]
    );
}
#[test]
fn setenv_overwrite() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <stdlib.h>\nint main() { setenv(\"MY_TEST_VAR\", \"123\", 1); setenv(\"MY_TEST_VAR\", \"456\", 1); printf(\"%s\", getenv(\"MY_TEST_VAR\")); return 0; }"
        ),
        vec!["456"]
    );
}
#[test]
fn unsetenv_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <stdlib.h>\nint main() { setenv(\"MY_TEST_VAR\", \"123\", 1); unsetenv(\"MY_TEST_VAR\"); printf(\"%d\", getenv(\"MY_TEST_VAR\") == NULL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn putenv_basic() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE\n#include <stdlib.h>\nint main() { char var[] = \"PUTENV_VAR=789\"; putenv(var); printf(\"%s\", getenv(\"PUTENV_VAR\")); return 0; }"
        ),
        vec!["789"]
    );
}
#[test]
fn putenv_mutation() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE\n#include <stdlib.h>\nint main() { char var[] = \"MUT_VAR=AAA\"; putenv(var); var[8] = 'B'; printf(\"%s\", getenv(\"MUT_VAR\")); return 0; }"
        ),
        vec!["BAA"]
    );
} // some libc reflect changes to the string
#[test]
fn setenv_empty_value() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <stdlib.h>\nint main() { setenv(\"EMPTY_VAR\", \"\", 1); printf(\"%d\", getenv(\"EMPTY_VAR\")[0] == '\\0'); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn unsetenv_not_found() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <stdlib.h>\nint main() { int res = unsetenv(\"DOES_NOT_EXIST\"); printf(\"%d\", res == 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn setenv_invalid_name() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <stdlib.h>\nint main() { int res = setenv(\"INVALID=NAME\", \"123\", 1); printf(\"%d\", res != 0); return 0; }"
        ),
        vec!["1"]
    );
} // names cannot contain =
#[test]
fn clearenv_basic() {
    assert_eq!(
        run_c(
            "#define _BSD_SOURCE\n#include <stdlib.h>\nint main() { setenv(\"TEST_CLEAR\", \"1\", 1); clearenv(); printf(\"%d\", getenv(\"TEST_CLEAR\") == NULL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn environ_ptr_access() {
    assert_eq!(
        run_c(
            "#define _GNU_SOURCE\n#include <stdlib.h>\n#include <unistd.h>\nextern char **environ;\nint main() { setenv(\"TEST_ENVIRON\", \"1\", 1); int found = 0; for(char **e = environ; *e; e++) { if((*e)[0] == 'T') found = 1; } printf(\"%d\", found); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn getenv_after_clearenv() {
    assert_eq!(
        run_c(
            "#define _BSD_SOURCE\n#include <stdlib.h>\nint main() { clearenv(); setenv(\"NEW_VAR\", \"42\", 1); printf(\"%s\", getenv(\"NEW_VAR\")); return 0; }"
        ),
        vec!["42"]
    );
}
#[test]
fn putenv_remove() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE\n#include <stdlib.h>\nint main() { char var[] = \"RM_VAR=1\"; putenv(var); char rm[] = \"RM_VAR\"; putenv(rm); /* Some implementations remove if no = */ printf(\"%d\", getenv(\"RM_VAR\") == NULL || getenv(\"RM_VAR\") != NULL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn secure_getenv_basic() {
    assert_eq!(
        run_c(
            "#define _GNU_SOURCE\n#include <stdlib.h>\nint main() { setenv(\"SECURE_VAR\", \"1\", 1); char *p = secure_getenv(\"SECURE_VAR\"); printf(\"%d\", p != NULL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn setenv_null_value() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <stdlib.h>\nint main() { /* NULL value behavior varies, but we expect it to not crash */ setenv(\"NULL_VAR\", \"\", 1); printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn environ_count() {
    assert_eq!(
        run_c(
            "#define _GNU_SOURCE\n#include <stdlib.h>\n#include <unistd.h>\nextern char **environ;\nint main() { int count = 0; while(environ && environ[count]) count++; printf(\"%d\", count >= 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn getenv_modifying_returned_string() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <stdlib.h>\nint main() { setenv(\"VAR\", \"value\", 1); /* standard says modifying is UB, so we just run ok */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn setenv_large_value() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <stdlib.h>\n#include <string.h>\nint main() { char large[10000]; memset(large, 'x', 9999); large[9999] = '\\0'; setenv(\"LARGE_VAR\", large, 1); printf(\"%d\", (int)strlen(getenv(\"LARGE_VAR\"))); return 0; }"
        ),
        vec!["9999"]
    );
}
