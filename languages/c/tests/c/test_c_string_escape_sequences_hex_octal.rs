use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn str_escape_newline() {
    assert_eq!(
        run_c("int main() { printf(\"%c\", '\\n'); return 0; }"),
        vec![""]
    );
}
#[test]
fn str_escape_tab() {
    assert_eq!(
        run_c("int main() { printf(\"%c\", '\\t'); return 0; }"),
        vec!["\t"]
    );
}
#[test]
fn str_escape_hex_basic() {
    assert_eq!(
        run_c("int main() { printf(\"%c\", '\\x41'); return 0; }"),
        vec!["A"]
    );
} // Hex 41 is 'A'
#[test]
fn str_escape_hex_uppercase() {
    assert_eq!(
        run_c("int main() { printf(\"%c\", '\\x4A'); return 0; }"),
        vec!["J"]
    );
}
#[test]
fn str_escape_hex_lowercase() {
    assert_eq!(
        run_c("int main() { printf(\"%c\", '\\x6a'); return 0; }"),
        vec!["j"]
    );
}
#[test]
fn str_escape_hex_multiple_digits() {
    assert_eq!(
        run_c("int main() { printf(\"%c\", '\\x0041'); return 0; }"),
        vec!["A"]
    );
} // Hex continues until non-hex
#[test]
fn str_escape_hex_limit_in_string() {
    assert_eq!(
        run_c("int main() { printf(\"%s\", \"\\x41B\"); return 0; }"),
        vec!["\u{1b}"]
    );
} // Hex escapes consume all following hex digits; char conversion keeps the low byte.
#[test]
fn str_escape_hex_limit_workaround() {
    assert_eq!(
        run_c("int main() { printf(\"%s\", \"\\x41\" \"B\"); return 0; }"),
        vec!["AB"]
    );
}
#[test]
fn str_escape_octal_basic() {
    assert_eq!(
        run_c("int main() { printf(\"%c\", '\\101'); return 0; }"),
        vec!["A"]
    );
} // Octal 101 is 65 is 'A'
#[test]
fn str_escape_octal_limit() {
    assert_eq!(
        run_c("int main() { printf(\"%s\", \"\\1012\"); return 0; }"),
        vec!["A2"]
    );
} // Octal is max 3 digits. So \101 is A, then 2.
#[test]
fn str_escape_octal_one_digit() {
    assert_eq!(
        run_c("int main() { printf(\"%c\", '\\0'); return 0; }"),
        vec!["\0"]
    );
}
#[test]
fn str_escape_octal_two_digits() {
    assert_eq!(
        run_c("int main() { printf(\"%c\", '\\12'); return 0; }"),
        vec![""]
    );
} // Octal 12 is newline; run_prints captures it as an empty line.
#[test]
fn str_escape_backslash() {
    assert_eq!(
        run_c("int main() { printf(\"%c\", '\\\\'); return 0; }"),
        vec!["\\"]
    );
}
#[test]
fn str_escape_single_quote() {
    assert_eq!(
        run_c("int main() { printf(\"%c\", '\\''); return 0; }"),
        vec!["'"]
    );
}
#[test]
fn str_escape_double_quote() {
    assert_eq!(
        run_c("int main() { printf(\"%s\", \"\\\"\"); return 0; }"),
        vec!["\""]
    );
}
#[test]
fn str_escape_question_mark() {
    assert_eq!(
        run_c("int main() { printf(\"%c\", '\\?'); return 0; }"),
        vec!["?"]
    );
}
#[test]
fn str_escape_universal_character_name() {
    assert_eq!(
        run_c("int main() { printf(\"%s\", \"\\u0041\"); return 0; }"),
        vec!["A"]
    );
}
