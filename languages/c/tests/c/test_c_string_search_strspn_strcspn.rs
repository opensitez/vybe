use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn strspn_basic() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { printf(\"%d\", (int)strspn(\"hello\", \"he\")); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn strspn_all_match() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { printf(\"%d\", (int)strspn(\"hello\", \"helo\")); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn strspn_no_match() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { printf(\"%d\", (int)strspn(\"hello\", \"xyz\")); return 0; }"
        ),
        vec!["0"]
    );
}
#[test]
fn strspn_empty_str() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { printf(\"%d\", (int)strspn(\"\", \"he\")); return 0; }"
        ),
        vec!["0"]
    );
}
#[test]
fn strspn_empty_accept() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { printf(\"%d\", (int)strspn(\"hello\", \"\")); return 0; }"
        ),
        vec!["0"]
    );
}
#[test]
fn strcspn_basic() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { printf(\"%d\", (int)strcspn(\"hello\", \"l\")); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn strcspn_all_match() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { printf(\"%d\", (int)strcspn(\"hello\", \"xyz\")); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn strcspn_first_char() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { printf(\"%d\", (int)strcspn(\"hello\", \"he\")); return 0; }"
        ),
        vec!["0"]
    );
}
#[test]
fn strcspn_empty_str() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { printf(\"%d\", (int)strcspn(\"\", \"l\")); return 0; }"
        ),
        vec!["0"]
    );
}
#[test]
fn strcspn_empty_reject() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { printf(\"%d\", (int)strcspn(\"hello\", \"\")); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn strpbrk_found() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { printf(\"%s\", strpbrk(\"hello\", \"l\")); return 0; }"
        ),
        vec!["llo"]
    );
}
#[test]
fn strpbrk_first_char() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { printf(\"%s\", strpbrk(\"hello\", \"he\")); return 0; }"
        ),
        vec!["hello"]
    );
}
#[test]
fn strpbrk_not_found() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { printf(\"%d\", strpbrk(\"hello\", \"xyz\") == NULL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn strpbrk_empty_str() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { printf(\"%d\", strpbrk(\"\", \"he\") == NULL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn strpbrk_empty_accept() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { printf(\"%d\", strpbrk(\"hello\", \"\") == NULL); return 0; }"
        ),
        vec!["1"]
    );
}
