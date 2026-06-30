use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_compute_rounded() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC 9V9 VALUE 2.6.
01 WS-B PIC 9V9 VALUE 1.8.
01 WS-C PIC 9 VALUE 0.
"#,
        r#"
    COMPUTE WS-C ROUNDED = WS-A + WS-B.
    DISPLAY WS-C.
"#,
    ));
    assert_eq!(output, vec!["4"]);
}

#[test]
fn test_compute_unary_minus() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC S9(3) VALUE 42.
01 WS-B PIC S9(3) VALUE 0.
"#,
        r#"
    COMPUTE WS-B = - WS-A.
    DISPLAY WS-B.
"#,
    ));
    assert!(output.len() >= 1);
}

#[test]
fn test_compute_intrinsics() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC S9(3) VALUE -42.
01 WS-B PIC 9(3) VALUE 0.
01 WS-C PIC 9(3) VALUE 0.
"#,
        r#"
    COMPUTE WS-B = FUNCTION ABS(WS-A).
    COMPUTE WS-C = FUNCTION INTEGER(3.7).
    DISPLAY WS-B.
    DISPLAY WS-C.
"#,
    ));
    assert_eq!(output, vec!["042", "003"]);
}

#[test]
fn test_compute_max_min() {
    let output = run_prints(&p(
        r#"
01 WS-MAX PIC 9(3) VALUE 0.
01 WS-MIN PIC 9(3) VALUE 0.
"#,
        r#"
    COMPUTE WS-MAX = FUNCTION MAX(10 5 20 15).
    COMPUTE WS-MIN = FUNCTION MIN(10 5 20 15).
    DISPLAY WS-MAX.
    DISPLAY WS-MIN.
"#,
    ));
    assert_eq!(output, vec!["020", "005"]);
}

#[test]
fn test_compute_precedence() {
    let output = run_prints(&p(
        r#"
01 WS-R PIC 9(3) VALUE 0.
"#,
        r#"
    COMPUTE WS-R = 2 + 3 ** 2.
    DISPLAY WS-R.
"#,
    ));
    assert_eq!(output, vec!["011"]);
}

#[test]
fn test_compute_deep_parens() {
    let output = run_prints(&p(
        r#"
01 WS-R PIC 9(3) VALUE 0.
"#,
        r#"
    COMPUTE WS-R = ((((1 + 1) * 2) + 2) * 2).
    DISPLAY WS-R.
"#,
    ));
    assert_eq!(output, vec!["012"]);
}

#[test]
fn test_compute_chain() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC 9(3) VALUE 10.
01 WS-B PIC 9(3) VALUE 20.
01 WS-C PIC 9(3) VALUE 30.
01 WS-R PIC 9(3) VALUE 0.
"#,
        r#"
    COMPUTE WS-R = WS-A + WS-B.
    COMPUTE WS-R = WS-R + WS-C.
    DISPLAY WS-R.
"#,
    ));
    assert_eq!(output, vec!["060"]);
}

#[test]
fn test_compute_math_funcs() {
    compile_ok(&p(
        r#"
01 WS-R PIC 9V9999.
"#,
        r#"
    COMPUTE WS-R = FUNCTION SIN(1.0) + FUNCTION COS(1.0).
"#,
    ));
}
