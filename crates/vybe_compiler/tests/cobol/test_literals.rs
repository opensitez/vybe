use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn numeric_and_alpha_literals_compile() {
    compile_ok(&p(
        "01 WS-NUM PIC 9(5) VALUE 100.\n01 WS-TXT PIC X(5) VALUE \"A\".",
        "    MOVE 12345 TO WS-NUM.\n    MOVE \"HELLO\" TO WS-TXT.",
    ));
}

#[test]
fn figurative_constants_compile() {
    compile_ok(&p(
        "01 WS-FILL PIC X(5) VALUE SPACES.\n01 WS-ZERO PIC 9(5) VALUE ZERO.",
        "    MOVE ZEROS TO WS-ZERO.\n    MOVE SPACES TO WS-FILL.",
    ));
}
