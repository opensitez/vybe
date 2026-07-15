use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn str_concat_basic() {
    assert_eq!(
        run_c("int main() { printf(\"%s\", \"Hello \" \"World\"); return 0; }"),
        vec!["Hello World"]
    );
}
#[test]
fn str_concat_three_parts() {
    assert_eq!(
        run_c("int main() { printf(\"%s\", \"A\" \"B\" \"C\"); return 0; }"),
        vec!["ABC"]
    );
}
#[test]
fn str_concat_with_newlines() {
    assert_eq!(
        run_c("int main() { printf(\"%s\", \"Hello\\n\" \"World\"); return 0; }"),
        vec!["Hello", "World"]
    );
}
#[test]
fn str_concat_with_spaces_between() {
    assert_eq!(
        run_c("int main() { printf(\"%s\", \"A\"   \"B\"); return 0; }"),
        vec!["AB"]
    );
}
#[test]
fn str_concat_in_array_init() {
    assert_eq!(
        run_c("int main() { char str[] = \"X\" \"Y\"; printf(\"%s\", str); return 0; }"),
        vec!["XY"]
    );
}
#[test]
fn str_concat_in_pointer_init() {
    assert_eq!(
        run_c("int main() { char *str = \"1\" \"2\"; printf(\"%s\", str); return 0; }"),
        vec!["12"]
    );
}
#[test]
fn str_concat_across_lines() {
    assert_eq!(
        run_c("int main() { printf(\"%s\", \"foo\" \n \"bar\"); return 0; }"),
        vec!["foobar"]
    );
}
#[test]
fn str_concat_with_escapes() {
    assert_eq!(
        run_c("int main() { printf(\"%s\", \"\\x41\" \"B\"); return 0; }"),
        vec!["AB"]
    );
} // A is 41
#[test]
fn str_concat_with_macro() {
    assert_eq!(
        run_c("#define STR \"bar\"\nint main() { printf(\"%s\", \"foo\" STR); return 0; }"),
        vec!["foobar"]
    );
}
#[test]
fn str_concat_empty() {
    assert_eq!(
        run_c("int main() { printf(\"%s\", \"\" \"A\"); return 0; }"),
        vec!["A"]
    );
}
#[test]
fn str_concat_multiple_empty() {
    assert_eq!(
        run_c("int main() { printf(\"%s\", \"\" \"\" \"B\"); return 0; }"),
        vec!["B"]
    );
}
#[test]
fn str_concat_sizeof() {
    assert_eq!(
        run_c("int main() { printf(\"%d\", (int)sizeof(\"A\" \"B\")); return 0; }"),
        vec!["3"]
    );
} // A, B, \0
#[test]
fn str_concat_char_pointer_array() {
    assert_eq!(
        run_c(
            "int main() { char *arr[] = {\"A\" \"B\", \"C\" \"D\"}; printf(\"%s%s\", arr[0], arr[1]); return 0; }"
        ),
        vec!["ABCD"]
    );
}
#[test]
fn str_concat_with_pragma() {
    assert_eq!(
        run_c(
            "int main() { _Pragma(\"GCC diagnostic ignored \\\"-Wpragmas\\\"\") printf(\"%s\", \"ok\" \"!\"); return 0; }"
        ),
        vec!["ok!"]
    );
}
