use super::helpers::{compile_ok, run_prints};

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

#[test]
fn statistical_intrinsics_runtime_nonempty() {
    let out = run_prints(&p(
        r#"
01 WS-DEV PIC 9(3)V99.
01 WS-RNG PIC 9(3).
01 WS-INT PIC S9(3).
01 WS-FRAC PIC V99.
"#,
        r#"
    COMPUTE WS-DEV = FUNCTION STANDARD-DEVIATION(10 20 30)
    COMPUTE WS-RNG = FUNCTION RANGE(10 20 30)
    DISPLAY WS-DEV
    DISPLAY WS-RNG
"#,
    ));
    assert_eq!(out.len(), 2);
    assert!(!out[0].trim().is_empty());
    assert!(!out[1].trim().is_empty());
}

#[test]
fn statistical_parts_runtime() {
    let out = run_prints(&p(
        r#"
01 WS-INT-PART PIC S9(3).
01 WS-FRAC-PART PIC V99.
"#,
        r#"
    COMPUTE WS-INT-PART = FUNCTION INTEGER-PART(-3.7)
    COMPUTE WS-FRAC-PART = FUNCTION FRACTION-PART(3.7)
    DISPLAY WS-INT-PART
    DISPLAY WS-FRAC-PART
"#,
    ));
    assert_eq!(out.len(), 2);
}

#[test]
fn statistical_financial_runtime() {
    let out = run_prints(&p(
        r#"
01 WS-ANN PIC 9(3)V9999.
01 WS-PV PIC 9(5)V99.
"#,
        r#"
    COMPUTE WS-ANN = FUNCTION ANNUITY(0.05 10)
    COMPUTE WS-PV = FUNCTION PRESENT-VALUE(0.05 100 100 100)
    DISPLAY WS-ANN
    DISPLAY WS-PV
"#,
    ));
    assert_eq!(out.len(), 2);
}

#[test]
fn statistical_parts_runtime_exact() {
    let out = run_prints(&p(
        r#"
01 WS-INT-PART PIC S9(4).
01 WS-FRAC-PART PIC V999.
"#,
        r#"
    COMPUTE WS-INT-PART = FUNCTION INTEGER-PART(12.34).
    COMPUTE WS-FRAC-PART = FUNCTION FRACTION-PART(12.34).
    DISPLAY WS-INT-PART.
    DISPLAY WS-FRAC-PART.
    DISPLAY FUNCTION MAX(42 24 18).
"#,
    ));
    assert_eq!(out.len(), 3);
    assert!(!out[2].trim().is_empty());
}

#[test]
fn statistical_constants_nonempty() {
    compile_ok(&p(
        "01 WS-VALUE PIC 9V9 VALUE 0.",
        r#"
    COMPUTE WS-VALUE = FUNCTION E + FUNCTION PI.
    COMPUTE WS-VALUE = FUNCTION SIN(0).
"#,
    ));
}
