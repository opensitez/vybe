use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn str_wide_literal_sizeof() {
    assert_eq!(
        run_c(
            "#include <stddef.h>\nint main() { printf(\"%d\", (int)(sizeof(L\"A\") / sizeof(wchar_t))); return 0; }"
        ),
        vec!["2"]
    );
} // A and null
#[test]
fn str_wide_literal_compile() {
    assert_eq!(
        run_c("int main() { const int *p = (const int *)L\"AB\"; printf(\"ok\"); return 0; }"),
        vec!["ok"]
    );
}
#[test]
fn str_wide_literal_concat() {
    assert_eq!(
        run_c(
            "int main() { const int *p = (const int *)(L\"A\" L\"B\"); printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn str_wide_literal_concat_mixed() {
    assert_eq!(
        run_c(
            "int main() { const int *p = (const int *)(L\"A\" \"B\"); printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
} // Result is wide
#[test]
fn str_utf8_literal_sizeof() {
    assert_eq!(
        run_c("int main() { printf(\"%d\", (int)sizeof(u8\"A\")); return 0; }"),
        vec!["2"]
    );
} // C11 u8 literal
#[test]
fn str_utf8_literal_concat() {
    assert_eq!(
        run_c("int main() { printf(\"%s\", u8\"A\" u8\"B\"); return 0; }"),
        vec!["AB"]
    );
}
#[test]
fn str_utf8_literal_concat_mixed() {
    assert_eq!(
        run_c("int main() { printf(\"%s\", u8\"A\" \"B\"); return 0; }"),
        vec!["AB"]
    );
}
#[test]
fn str_utf16_literal_sizeof() {
    assert_eq!(
        run_c(
            "#include <uchar.h>\nint main() { printf(\"%d\", (int)(sizeof(u\"A\") / sizeof(char16_t))); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn str_utf16_literal_concat() {
    assert_eq!(
        run_c(
            "#include <uchar.h>\nint main() { const char16_t *p = u\"A\" u\"B\"; printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn str_utf32_literal_sizeof() {
    assert_eq!(
        run_c(
            "#include <uchar.h>\nint main() { printf(\"%d\", (int)(sizeof(U\"A\") / sizeof(char32_t))); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn str_utf32_literal_concat() {
    assert_eq!(
        run_c(
            "#include <uchar.h>\nint main() { const char32_t *p = U\"A\" U\"B\"; printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn str_utf8_multibyte() {
    assert_eq!(
        run_c(
            "int main() { char str[] = u8\"\\u00A9\"; printf(\"%d\", (int)sizeof(str)); return 0; }"
        ),
        vec!["3"]
    );
} // copyright symbol is 2 bytes + null
#[test]
fn str_wide_multibyte() {
    assert_eq!(
        run_c(
            "#include <stddef.h>\nint main() { wchar_t str[] = L\"\\u00A9\"; printf(\"%d\", (int)(sizeof(str)/sizeof(wchar_t))); return 0; }"
        ),
        vec!["2"]
    );
} // 1 wchar + null
#[test]
fn str_utf16_multibyte() {
    assert_eq!(
        run_c(
            "#include <uchar.h>\nint main() { char16_t str[] = u\"\\U0001F600\"; printf(\"%d\", (int)(sizeof(str)/sizeof(char16_t))); return 0; }"
        ),
        vec!["3"]
    );
} // Emoji requires surrogate pair + null
