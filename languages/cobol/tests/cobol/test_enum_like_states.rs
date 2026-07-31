use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn condition_name_state_set_true_compiles() {
    compile_ok(&p(
        "01 WS-STATE PIC 9 VALUE 0.\n   88 STATE-NEW VALUE 1.\n   88 STATE-DONE VALUE 2.",
        "    SET STATE-NEW TO TRUE.\n    IF STATE-NEW\n        DISPLAY \"NEW\"\n    END-IF.",
    ));
}

#[test]
fn condition_name_state_transition_compiles() {
    compile_ok(&p(
        "01 WS-STATE PIC 9 VALUE 0.\n   88 STATE-NEW VALUE 1.\n   88 STATE-RUN VALUE 2.\n   88 STATE-DONE VALUE 3.",
        "    SET STATE-NEW TO TRUE.\n    SET STATE-RUN TO TRUE.\n    SET STATE-DONE TO TRUE.\n    IF STATE-DONE\n        DISPLAY \"DONE\"\n    END-IF.",
    ));
}

#[test]
fn condition_name_false_assignment_compiles() {
    compile_ok(&p(
        "01 WS-FLAG PIC 9 VALUE 1.\n   88 IS-ON VALUE 1.\n   88 IS-OFF VALUE 0.",
        "    SET IS-OFF TO TRUE.\n    IF IS-OFF DISPLAY \"OFF\" END-IF.",
    ));
}

#[test]
fn evaluate_on_enum_like_state_compiles() {
    compile_ok(&p(
        "01 WS-STATE PIC 9 VALUE 2.",
        "    EVALUATE WS-STATE\n        WHEN 1 DISPLAY \"NEW\"\n        WHEN 2 DISPLAY \"RUN\"\n        WHEN 3 DISPLAY \"DONE\"\n        WHEN OTHER DISPLAY \"UNK\"\n    END-EVALUATE.",
    ));
}

#[test]
fn enum_state_transition_runtime() {
    let out = run_prints(&p(
        "01 WS-STATE PIC 9 VALUE 0.\n   88 STATE-NEW VALUE 1.\n   88 STATE-RUN VALUE 2.\n   88 STATE-DONE VALUE 3.",
        "    SET STATE-NEW TO TRUE\n    IF STATE-NEW\n        DISPLAY \"NEW\"\n    END-IF\n    SET STATE-RUN TO TRUE\n    IF STATE-RUN\n        DISPLAY \"RUN\"\n    END-IF",
    ));
    assert_eq!(out, vec!["NEW", "RUN"]);
}

#[test]
fn enum_state_invalid_path_runtime() {
    let out = run_prints(&p(
        "01 WS-STATE PIC 9 VALUE 9.\n   88 STATE-NEW VALUE 1.\n   88 STATE-DONE VALUE 2.",
        "    IF NOT STATE-NEW\n       AND NOT STATE-DONE\n        DISPLAY \"UNKNOWN\"\n    END-IF",
    ));
    assert_eq!(out, vec!["UNKNOWN"]);
}

#[test]
fn enum_state_transition_to_false_surface() {
    let out = run_prints(&p(
        "01 WS-STATE PIC 9 VALUE 2.\n   88 STATE-NEW VALUE 1.\n   88 STATE-BUSY VALUE 2.\n   88 STATE-DONE VALUE 3.",
        "    SET STATE-BUSY TO TRUE\n    DISPLAY WS-STATE\n    SET STATE-BUSY TO FALSE\n    IF STATE-BUSY\n        DISPLAY \"BUSY\"\n    ELSE\n        DISPLAY \"NOT-BUSY\"\n    END-IF",
    ));
    assert_eq!(out, vec!["2", "NOT-BUSY"]);
}
