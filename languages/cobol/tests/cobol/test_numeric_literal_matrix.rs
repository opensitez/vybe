use super::helpers::compile_ok;

fn build_program(value: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-N PIC 9(3) VALUE {value}.\nPROCEDURE DIVISION.\n    IF WS-N > 0\n        DISPLAY \"OK\"\n    END-IF.\n    STOP RUN."
    )
}

#[test]
fn numeric_literal_zero_value_compiles() {
    compile_ok(&build_program("0"));
}

#[test]
fn numeric_literal_positive_value_compiles() {
    compile_ok(&build_program("123"));
}

#[test]
fn numeric_literal_negative_value_compiles() {
    compile_ok(&build_program("-7"));
}

#[test]
fn numeric_literal_leading_zero_value_compiles() {
    compile_ok(&build_program("007"));
}

#[test]
fn numeric_literal_large_value_compiles() {
    compile_ok(&build_program("999"));
}

#[test]
fn numeric_literal_boundary_value_compiles() {
    compile_ok(&build_program("001"));
}
