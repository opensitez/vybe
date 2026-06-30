use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_date_conversions() {
    compile_ok(&p(
        r#"
01 WS-INT PIC 9(8) VALUE 20240101.
01 WS-DAY PIC 9(7) VALUE 2024001.
01 WS-JULIAN PIC 9(8).
"#,
        r#"
    COMPUTE WS-JULIAN = FUNCTION INTEGER-OF-DATE(WS-INT).
    COMPUTE WS-JULIAN = FUNCTION INTEGER-OF-DAY(WS-DAY).
    COMPUTE WS-INT = FUNCTION DATE-OF-INTEGER(WS-JULIAN).
    COMPUTE WS-DAY = FUNCTION DAY-OF-INTEGER(WS-JULIAN).
"#,
    ));
}

#[test]
fn test_year_to_yyyy() {
    compile_ok(&p(
        "01 WS-YEAR PIC 9(4).",
        r#"
    COMPUTE WS-YEAR = FUNCTION YEAR-TO-YYYY(24 50 2000).
"#,
    ));
}

#[test]
fn test_date_validations() {
    compile_ok(&p(
        "01 WS-RES PIC 9(9).",
        r#"
    COMPUTE WS-RES = FUNCTION TEST-DATE-YYYYMMDD(20240229).
    COMPUTE WS-RES = FUNCTION TEST-DAY-YYYYDDD(2024001).
"#,
    ));
}

#[test]
fn test_time_intrinsics() {
    compile_ok(&p(
        "01 WS-SEC PIC 9(9).",
        r#"
    COMPUTE WS-SEC = FUNCTION SECONDS-FROM-MIDNIGHT.
"#,
    ));
}
