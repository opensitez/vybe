use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn scanf_suppress_int() {
    assert_eq!(
        run_c(
            "int main() { int a = 0; int n = sscanf(\"123 456\", \"%*d %d\", &a); printf(\"%d %d\", n, a); return 0; }"
        ),
        vec!["1 456"]
    );
} // Returns 1 assignment
#[test]
fn scanf_suppress_string() {
    assert_eq!(
        run_c(
            "int main() { char buf[10]; int n = sscanf(\"hello world\", \"%*s %s\", buf); printf(\"%d %s\", n, buf); return 0; }"
        ),
        vec!["1 world"]
    );
}
#[test]
fn scanf_suppress_char() {
    assert_eq!(
        run_c(
            "int main() { char c = 'x'; int n = sscanf(\"abc\", \"%*c%c\", &c); printf(\"%d %c\", n, c); return 0; }"
        ),
        vec!["1 b"]
    );
}
#[test]
fn scanf_suppress_scanset() {
    assert_eq!(
        run_c(
            "int main() { char buf[10]; int n = sscanf(\"abc123def\", \"%*[a-z]%[0-9]\", buf); printf(\"%d %s\", n, buf); return 0; }"
        ),
        vec!["1 123"]
    );
}
#[test]
fn scanf_suppress_float() {
    assert_eq!(
        run_c(
            "int main() { float f; int n = sscanf(\"1.23 4.56\", \"%*f %f\", &f); printf(\"%d %.2f\", n, f); return 0; }"
        ),
        vec!["1 4.56"]
    );
}
#[test]
fn scanf_suppress_all() {
    assert_eq!(
        run_c(
            "int main() { int n = sscanf(\"123 abc\", \"%*d %*s\"); printf(\"%d\", n); return 0; }"
        ),
        vec!["0"]
    );
}
#[test]
fn scanf_suppress_width_limit() {
    assert_eq!(
        run_c(
            "int main() { int a; int n = sscanf(\"123456\", \"%*3d%d\", &a); printf(\"%d %d\", n, a); return 0; }"
        ),
        vec!["1 456"]
    );
}
#[test]
fn scanf_match_literal_chars() {
    assert_eq!(
        run_c(
            "int main() { int a, b; int n = sscanf(\"123-456\", \"%d-%d\", &a, &b); printf(\"%d %d %d\", n, a, b); return 0; }"
        ),
        vec!["2 123 456"]
    );
}
#[test]
fn scanf_match_literal_spaces() {
    assert_eq!(
        run_c(
            "int main() { int a, b; int n = sscanf(\"123   456\", \"%d %d\", &a, &b); printf(\"%d %d %d\", n, a, b); return 0; }"
        ),
        vec!["2 123 456"]
    );
} // space in fmt matches any amount of whitespace
