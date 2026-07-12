use super::helpers::run_prints;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn string_move_literal_and_variable_preserves_value() {
    let output = run_prints(&p(
        "01 WS-A PIC X(8) VALUE \"ALPHA\".\n01 WS-B PIC X(8) VALUE SPACES.",
        "    MOVE WS-A TO WS-B.\n    DISPLAY WS-B.",
    ));
    assert_eq!(output, vec!["ALPHA"]);
}

#[test]
fn string_concat_via_string_statement_produces_joined_text() {
    let output = run_prints(&p(
        "01 WS-A PIC X(5) VALUE \"HEL\".\n01 WS-B PIC X(5) VALUE \"LO\".\n01 WS-C PIC X(10) VALUE SPACES.",
        "    STRING WS-A DELIMITED BY SIZE WS-B DELIMITED BY SIZE INTO WS-C.\n    DISPLAY WS-C.",
    ));
    assert_eq!(output, vec!["HELLO"]);
}

#[test]
fn unstring_basic_assigns_each_target_field() {
    let output = run_prints(&p(
        "01 WS-SRC PIC X(12) VALUE \"A,B,C\".\n01 WS-A PIC X(3) VALUE SPACES.\n01 WS-B PIC X(3) VALUE SPACES.\n01 WS-C PIC X(3) VALUE SPACES.",
        "    UNSTRING WS-SRC DELIMITED BY \",\" INTO WS-A WS-B WS-C.\n    DISPLAY WS-A.\n    DISPLAY WS-B.\n    DISPLAY WS-C.",
    ));
    assert_eq!(output, vec!["A", "B", "C"]);
}

#[test]
fn inspect_tallying_counts_all_matches() {
    let output = run_prints(&p(
        "01 WS-TXT PIC X(10) VALUE \"ABCA\".\n01 WS-COUNT PIC 9(4) VALUE 0.",
        "    INSPECT WS-TXT TALLYING WS-COUNT FOR ALL \"A\".\n    DISPLAY WS-COUNT.",
    ));
    assert_eq!(output, vec!["2"]);
}

#[test]
fn inspect_replacing_rewrites_requested_characters() {
    let output = run_prints(&p(
        "01 WS-TXT PIC X(10) VALUE \"ABCA\".",
        "    INSPECT WS-TXT REPLACING FIRST \"A\" BY \"Z\".\n    DISPLAY WS-TXT.",
    ));
    assert_eq!(output, vec!["ZBCA"]);
}

#[test]
fn reference_modification_extracts_text_slice() {
    let output = run_prints(&p(
        "01 WS-TXT PIC X(10) VALUE \"HELLOTEST\".\n01 WS-SUB PIC X(4) VALUE SPACES.",
        "    MOVE WS-TXT(6:4) TO WS-SUB.\n    DISPLAY WS-SUB.",
    ));
    assert_eq!(output, vec!["TEST"]);
}
