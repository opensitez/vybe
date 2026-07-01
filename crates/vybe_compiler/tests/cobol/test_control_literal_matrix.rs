use super::helpers::compile_ok;

fn build_program(value: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-N PIC 9(1) VALUE {value}.\nPROCEDURE DIVISION.\n    EVALUATE WS-N\n        WHEN 0\n            DISPLAY \"ZERO\"\n        WHEN 1\n            DISPLAY \"ONE\"\n        WHEN OTHER\n            DISPLAY \"OTHER\"\n    END-EVALUATE.\n    STOP RUN."
    )
}

#[test]
fn control_evaluate_zero_branch_compiles() {
    compile_ok(&build_program("0"));
}

#[test]
fn control_evaluate_one_branch_compiles() {
    compile_ok(&build_program("1"));
}

#[test]
fn control_evaluate_other_branch_compiles() {
    compile_ok(&build_program("2"));
}

#[test]
fn control_evaluate_with_negative_value_compiles() {
    compile_ok(&build_program("-1"));
}

#[test]
fn control_evaluate_with_large_value_compiles() {
    compile_ok(&build_program("9"));
}

#[test]
fn control_evaluate_with_boundary_value_compiles() {
    compile_ok(&build_program("8"));
}
