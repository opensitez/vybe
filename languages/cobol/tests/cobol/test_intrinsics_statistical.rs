use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_intrinsics_statistical_advanced() {
    compile_ok(&p(
        r#"
01 WS-DEV PIC 9(3)V99.
01 WS-RNG PIC 9(3).
01 WS-VAL PIC S9(3)V99.
01 WS-INT-PART PIC S9(3).
01 WS-FRAC-PART PIC V99.
"#,
        r#"
    COMPUTE WS-DEV = FUNCTION STANDARD-DEVIATION(10 20 30).
    COMPUTE WS-RNG = FUNCTION RANGE(10 20 30).
    COMPUTE WS-INT-PART = FUNCTION INTEGER-PART(-3.7).
    COMPUTE WS-FRAC-PART = FUNCTION FRACTION-PART(3.7).
"#,
    ));
}

#[test]
fn test_intrinsics_financial() {
    compile_ok(&p(
        r#"
01 WS-ANN PIC 9(3)V9999.
01 WS-PV PIC 9(5)V99.
"#,
        r#"
    COMPUTE WS-ANN = FUNCTION ANNUITY(0.05 10).
    COMPUTE WS-PV = FUNCTION PRESENT-VALUE(0.05 100 100 100).
"#,
    ));
}

#[test]
fn test_intrinsics_constants() {
    compile_ok(&p(
        "01 WS-VAL PIC 9V999999.",
        r#"
    COMPUTE WS-VAL = FUNCTION PI + FUNCTION E.
"#,
    ));
}
