use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn printf_percent_n() {
    assert_eq!(
        run_c("int main() { int n = 0; printf(\"123%n\", &n); printf(\" %d\", n); return 0; }"),
        vec!["123 3"]
    );
}
#[test]
fn printf_width_star_negative() {
    assert_eq!(
        run_c("int main() { printf(\"|%*s|\", -5, \"x\"); return 0; }"),
        vec!["|x    |"]
    );
} // negative width means left justify
#[test]
fn printf_precision_star_negative() {
    assert_eq!(
        run_c("int main() { printf(\"|%.*f|\", -1, 3.14); return 0; }"),
        vec!["|3.140000|"]
    );
} // negative precision is ignored (treated as missing)
#[test]
fn printf_plus_flag_zero() {
    assert_eq!(
        run_c("int main() { printf(\"%+d\", 0); return 0; }"),
        vec!["+0"]
    );
}
#[test]
fn printf_space_flag_zero() {
    assert_eq!(
        run_c("int main() { printf(\"% d\", 0); return 0; }"),
        vec![" 0"]
    );
}
#[test]
fn printf_hash_flag_octal_zero() {
    assert_eq!(
        run_c("int main() { printf(\"%#o\", 0); return 0; }"),
        vec!["0"]
    );
}
#[test]
fn printf_hash_flag_hex_zero() {
    assert_eq!(
        run_c("int main() { printf(\"%#x\", 0); return 0; }"),
        vec!["0"]
    );
} // 0x is not prepended for 0
#[test]
fn printf_hash_flag_float() {
    assert_eq!(
        run_c("int main() { printf(\"%#.0f\", 3.0); return 0; }"),
        vec!["3."]
    );
} // forces decimal point
#[test]
fn scanf_suppression_assignment() {
    assert_eq!(
        run_c(
            "int main() { int x = 0; sscanf(\"10 20 30\", \"%*d %d %*d\", &x); printf(\"%d\", x); return 0; }"
        ),
        vec!["20"]
    );
}
#[test]
fn scanf_character_scanset_caret() {
    assert_eq!(
        run_c(
            "int main() { char buf[10] = {0}; sscanf(\"abc123def\", \"%[^0-9]\", buf); printf(\"%s\", buf); return 0; }"
        ),
        vec!["abc"]
    );
}
#[test]
fn scanf_bracket_dash() {
    assert_eq!(
        run_c(
            "int main() { char buf[10] = {0}; sscanf(\"a-b-c\", \"%[a-c-]\", buf); printf(\"%s\", buf); return 0; }"
        ),
        vec!["a-b-c"]
    );
}
#[test]
fn scanf_bracket_caret_dash() {
    assert_eq!(
        run_c(
            "int main() { char buf[10] = {0}; sscanf(\"12-34\", \"%[^a-z]\", buf); printf(\"%s\", buf); return 0; }"
        ),
        vec!["12-34"]
    );
}
#[test]
fn printf_pointer_null() {
    assert_eq!(
        run_c(
            "int main() { /* %p with NULL often prints (nil) or 0x0, we just test it doesn't crash */ char buf[20]; sprintf(buf, \"%p\", NULL); printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn scanf_percent_n() {
    assert_eq!(
        run_c(
            "int main() { int x, n = 0; sscanf(\"123 456\", \"%d %n\", &x, &n); printf(\"%d\", n); return 0; }"
        ),
        vec!["4"]
    );
} // %n does not count as a match for the return value
#[test]
fn scanf_eof_immediate() {
    assert_eq!(
        run_c(
            "int main() { int x=9; int res = sscanf(\"\", \"%d\", &x); printf(\"%d %d\", res, x); return 0; }"
        ),
        vec!["-1 9"]
    );
} // EOF = -1
#[test]
fn printf_ll_modifier() {
    assert_eq!(
        run_c("int main() { printf(\"%lld\", 9223372036854775807LL); return 0; }"),
        vec!["9223372036854775807"]
    );
}
#[test]
fn printf_hh_modifier() {
    assert_eq!(
        run_c("int main() { char c = 127; printf(\"%hhd\", c); return 0; }"),
        vec!["127"]
    );
}
#[test]
fn printf_z_modifier() {
    assert_eq!(
        run_c("#include <stddef.h>\nint main() { size_t s = 42; printf(\"%zu\", s); return 0; }"),
        vec!["42"]
    );
}
#[test]
fn printf_t_modifier() {
    assert_eq!(
        run_c(
            "#include <stddef.h>\nint main() { ptrdiff_t p = -42; printf(\"%td\", p); return 0; }"
        ),
        vec!["-42"]
    );
}
#[test]
fn printf_zero_pad_with_precision() {
    assert_eq!(
        run_c("int main() { printf(\"%05.3d\", 12); return 0; }"),
        vec!["  012"]
    );
} // zero flag ignored when precision is given for integers
