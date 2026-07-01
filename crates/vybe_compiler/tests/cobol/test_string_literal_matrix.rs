use super::helpers::compile_ok;

fn build_program(value: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-S PIC X(20) VALUE \"{value}\".\nPROCEDURE DIVISION.\n    STRING WS-S DELIMITED BY SIZE INTO WS-S.\n    STOP RUN."
    )
}

#[test]
fn string_literal_simple_word_compiles() {
    compile_ok(&build_program("HELLO"));
}

#[test]
fn string_literal_with_spaces_compiles() {
    compile_ok(&build_program("HELLO WORLD"));
}

#[test]
fn string_literal_with_punctuation_compiles() {
    compile_ok(&build_program("A,B.C"));
}

#[test]
fn string_literal_with_digits_compiles() {
    compile_ok(&build_program("12345"));
}

#[test]
fn string_literal_with_quotes_compiles() {
    compile_ok(&build_program("A\"B"));
}

#[test]
fn string_literal_with_mixed_case_compiles() {
    compile_ok(&build_program("MiXeD"));
}

#[test]
fn string_literal_with_trailing_spaces_compiles() {
    compile_ok(&build_program("END   "));
}
