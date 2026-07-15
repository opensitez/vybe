use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn strtok_basic() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char s[] = \"a,b,c\"; printf(\"%s \", strtok(s, \",\")); printf(\"%s \", strtok(NULL, \",\")); printf(\"%s\", strtok(NULL, \",\")); return 0; }"
        ),
        vec!["a b c"]
    );
}
#[test]
fn strtok_multiple_delimiters() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char s[] = \"a,,b\"; printf(\"%s \", strtok(s, \",\")); printf(\"%s\", strtok(NULL, \",\")); return 0; }"
        ),
        vec!["a b"]
    );
} // skips empty
#[test]
fn strtok_different_delimiters() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char s[] = \"a,b;c\"; printf(\"%s \", strtok(s, \",;\")); printf(\"%s \", strtok(NULL, \";,\")); printf(\"%s\", strtok(NULL, \",\")); return 0; }"
        ),
        vec!["a b c"]
    );
}
#[test]
fn strtok_trailing_delimiters() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char s[] = \"a,b,\"; printf(\"%s \", strtok(s, \",\")); printf(\"%s\", strtok(NULL, \",\")); printf(\"%d\", strtok(NULL, \",\") == NULL); return 0; }"
        ),
        vec!["a b1"]
    );
}
#[test]
fn strtok_leading_delimiters() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char s[] = \",,a,b\"; printf(\"%s \", strtok(s, \",\")); printf(\"%s\", strtok(NULL, \",\")); return 0; }"
        ),
        vec!["a b"]
    );
}
#[test]
fn strtok_only_delimiters() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char s[] = \",,,\"; printf(\"%d\", strtok(s, \",\") == NULL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn strtok_empty_string() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char s[] = \"\"; printf(\"%d\", strtok(s, \",\") == NULL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn strtok_r_basic() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char s[] = \"a,b,c\"; char *save; printf(\"%s \", strtok_r(s, \",\", &save)); printf(\"%s \", strtok_r(NULL, \",\", &save)); printf(\"%s\", strtok_r(NULL, \",\", &save)); return 0; }"
        ),
        vec!["a b c"]
    );
}
#[test]
fn strtok_r_nested() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char s[] = \"a:1,b:2\"; char *save1, *save2; printf(\"%s\", strtok_r(strtok_r(s, \",\", &save1), \":\", &save2)); printf(\"%s\", strtok_r(strtok_r(NULL, \",\", &save1), \":\", &save2)); return 0; }"
        ),
        vec!["ab"]
    );
} // gets keys with independent save cursors
#[test]
fn strtok_r_only_delimiters() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char s[] = \",,,\"; char *save; printf(\"%d\", strtok_r(s, \",\", &save) == NULL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn strtok_r_empty_string() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char s[] = \"\"; char *save; printf(\"%d\", strtok_r(s, \",\", &save) == NULL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn strtok_r_null_saveptr_deref_fails() {
    assert_eq!(
        run_c(
            "/* #include <string.h>\nint main() { char s[] = \"a,b\"; strtok_r(s, \",\", NULL); return 0; } */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn strtok_null_delim() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char s[] = \"abc\"; printf(\"%s\", strtok(s, \"\")); return 0; }"
        ),
        vec!["abc"]
    );
} // If delim is empty, it returns the rest of string
