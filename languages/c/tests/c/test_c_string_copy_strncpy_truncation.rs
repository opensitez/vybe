use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn strncpy_basic() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char dest[10]; strncpy(dest, \"hello\", 10); printf(\"%s\", dest); return 0; }"
        ),
        vec!["hello"]
    );
}
#[test]
fn strncpy_truncation() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char dest[4]; strncpy(dest, \"hello\", 3); dest[3] = '\\0'; printf(\"%s\", dest); return 0; }"
        ),
        vec!["hel"]
    );
}
#[test]
fn strncpy_padding_nulls() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char dest[6] = \"xxxxx\"; strncpy(dest, \"hi\", 5); printf(\"%d %d %d\", dest[2], dest[3], dest[4]); return 0; }"
        ),
        vec!["0 0 0"]
    );
}
#[test]
fn strncpy_no_null_term_if_long() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char dest[5] = \"xxxxx\"; strncpy(dest, \"hello\", 5); printf(\"%c%c%c%c%c\", dest[0], dest[1], dest[2], dest[3], dest[4]); return 0; }"
        ),
        vec!["hello"]
    );
}
#[test]
fn strncat_basic() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char dest[10] = \"hi\"; strncat(dest, \" there\", 10 - strlen(dest) - 1); printf(\"%s\", dest); return 0; }"
        ),
        vec!["hi there"]
    );
}
#[test]
fn strncat_truncation() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char dest[10] = \"hello \"; strncat(dest, \"world\", 3); printf(\"%s\", dest); return 0; }"
        ),
        vec!["hello wor"]
    );
}
#[test]
fn strncat_always_null_terminates() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char dest[4] = \"a\"; strncat(dest, \"bcdef\", 2); printf(\"%s\", dest); return 0; }"
        ),
        vec!["abc"]
    );
} // len becomes 1 + 2 = 3, plus null
#[test]
fn strncat_zero_n() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char dest[10] = \"abc\"; strncat(dest, \"def\", 0); printf(\"%s\", dest); return 0; }"
        ),
        vec!["abc"]
    );
}
#[test]
fn memmove_overlap_fwd() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char s[] = \"abcdef\"; memmove(s + 2, s, 4); printf(\"%s\", s); return 0; }"
        ),
        vec!["ababcd"]
    );
}
#[test]
fn memmove_overlap_bwd() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char s[] = \"abcdef\"; memmove(s, s + 2, 4); printf(\"%s\", s); return 0; }"
        ),
        vec!["cdefef"]
    );
}
#[test]
fn memcpy_basic() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char s1[] = \"abc\"; char s2[4]; memcpy(s2, s1, 4); printf(\"%s\", s2); return 0; }"
        ),
        vec!["abc"]
    );
}
#[test]
fn memccpy_found() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char dest[10]; char *p = memccpy(dest, \"hello\", 'l', 5); *p = '\\0'; printf(\"%s\", dest); return 0; }"
        ),
        vec!["hel"]
    );
}
#[test]
fn memccpy_not_found() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char dest[10]; char *p = memccpy(dest, \"hello\", 'x', 5); printf(\"%d\", p == NULL); return 0; }"
        ),
        vec!["1"]
    );
}
