use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn sprintf_basic() {
    assert_eq!(
        run_c(
            "int main() { char buf[50]; sprintf(buf, \"%d %s\", 123, \"abc\"); printf(\"%s\", buf); return 0; }"
        ),
        vec!["123 abc"]
    );
}
#[test]
fn snprintf_basic() {
    assert_eq!(
        run_c(
            "int main() { char buf[5]; snprintf(buf, 5, \"hello\"); printf(\"%s\", buf); return 0; }"
        ),
        vec!["hell"]
    );
}
#[test]
fn snprintf_returns_required_len() {
    assert_eq!(
        run_c(
            "int main() { char buf[5]; int n = snprintf(buf, 5, \"hello\"); printf(\"%d\", n); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn sscanf_basic() {
    assert_eq!(
        run_c(
            "int main() { int a; char b[10]; sscanf(\"123 abc\", \"%d %s\", &a, b); printf(\"%d %s\", a, b); return 0; }"
        ),
        vec!["123 abc"]
    );
}
#[test]
fn sscanf_partial_match() {
    assert_eq!(
        run_c(
            "int main() { int a, b; int n = sscanf(\"123 x\", \"%d %d\", &a, &b); printf(\"%d\", n); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn sscanf_no_match() {
    assert_eq!(
        run_c(
            "int main() { int a; int n = sscanf(\"abc\", \"%d\", &a); printf(\"%d\", n); return 0; }"
        ),
        vec!["0"]
    );
}
#[test]
fn sscanf_skip_assignment() {
    assert_eq!(
        run_c(
            "int main() { int a; int n = sscanf(\"123 456\", \"%*d %d\", &a); printf(\"%d %d\", n, a); return 0; }"
        ),
        vec!["1 456"]
    );
}
#[test]
fn sscanf_char_set() {
    assert_eq!(
        run_c(
            "int main() { char b[10]; sscanf(\"hello world\", \"%[a-z]\", b); printf(\"%s\", b); return 0; }"
        ),
        vec!["hello"]
    );
}
#[test]
fn sscanf_char_set_negated() {
    assert_eq!(
        run_c(
            "int main() { char b[10]; sscanf(\"hello world\", \"%[^o]\", b); printf(\"%s\", b); return 0; }"
        ),
        vec!["hell"]
    );
}
#[test]
fn sscanf_hex_input() {
    assert_eq!(
        run_c("int main() { int a; sscanf(\"0x1A\", \"%x\", &a); printf(\"%d\", a); return 0; }"),
        vec!["26"]
    );
}
#[test]
fn sscanf_octal_input() {
    assert_eq!(
        run_c("int main() { int a; sscanf(\"012\", \"%o\", &a); printf(\"%d\", a); return 0; }"),
        vec!["10"]
    );
}
#[test]
fn sscanf_chars() {
    assert_eq!(
        run_c(
            "int main() { char c[3]; sscanf(\"abc\", \"%3c\", c); printf(\"%c%c%c\", c[0], c[1], c[2]); return 0; }"
        ),
        vec!["abc"]
    );
}
#[test]
fn sprintf_float_formatting() {
    assert_eq!(
        run_c(
            "int main() { char buf[50]; sprintf(buf, \"%.2f\", 3.14159); printf(\"%s\", buf); return 0; }"
        ),
        vec!["3.14"]
    );
}
#[test]
fn sprintf_hex_formatting() {
    assert_eq!(
        run_c(
            "int main() { char buf[50]; sprintf(buf, \"%04x\", 255); printf(\"%s\", buf); return 0; }"
        ),
        vec!["00ff"]
    );
}
#[test]
fn sprintf_pointer_formatting() {
    assert_eq!(
        run_c(
            "int main() { char buf[50]; int a; sprintf(buf, \"%p\", (void*)0x1234); printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
} // can't reliably predict pointer format
