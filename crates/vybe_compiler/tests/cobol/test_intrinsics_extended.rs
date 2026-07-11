use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn intrinsic_upper_case_compiles() {
    compile_ok(&p(
        "01 WS-TXT PIC X(10) VALUE \"abc\".\n01 WS-OUT PIC X(10).",
        "    MOVE FUNCTION UPPER-CASE(WS-TXT) TO WS-OUT.",
    ));
}

#[test]
fn intrinsic_lower_case_compiles() {
    compile_ok(&p(
        "01 WS-TXT PIC X(10) VALUE \"ABC\".\n01 WS-OUT PIC X(10).",
        "    MOVE FUNCTION LOWER-CASE(WS-TXT) TO WS-OUT.",
    ));
}

#[test]
fn intrinsic_length_compiles() {
    compile_ok(&p(
        "01 WS-TXT PIC X(10) VALUE \"HELLO\".\n01 WS-LEN PIC 9(3).",
        "    MOVE FUNCTION LENGTH(WS-TXT) TO WS-LEN.",
    ));
}

#[test]
fn intrinsic_trim_compiles() {
    compile_ok(&p(
        "01 WS-TXT PIC X(10) VALUE \"  HI  \".\n01 WS-OUT PIC X(10).",
        "    MOVE FUNCTION TRIM(WS-TXT) TO WS-OUT.",
    ));
}

#[test]
fn intrinsic_reverse_compiles() {
    compile_ok(&p(
        "01 WS-TXT PIC X(10) VALUE \"HELLO\".\n01 WS-OUT PIC X(10).",
        "    MOVE FUNCTION REVERSE(WS-TXT) TO WS-OUT.",
    ));
}

#[test]
fn intrinsic_substitute_compiles() {
    compile_ok(&p(
        "01 WS-TXT PIC X(20) VALUE \"HELLO WORLD\".\n01 WS-OUT PIC X(20).",
        "    MOVE FUNCTION SUBSTITUTE(WS-TXT \"WORLD\" \"COBOL\") TO WS-OUT.",
    ));
}
