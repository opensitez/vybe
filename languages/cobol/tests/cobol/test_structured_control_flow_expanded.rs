use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn nested_if_evaluate_mix_compiles() {
    compile_ok(&p(
        "01 WS-X PIC 9 VALUE 2.",
        "    IF WS-X > 0\n        EVALUATE WS-X\n            WHEN 1 DISPLAY \"ONE\"\n            WHEN 2 DISPLAY \"TWO\"\n            WHEN OTHER DISPLAY \"OTHER\"\n        END-EVALUATE\n    END-IF.",
    ));
}

#[test]
fn perform_thru_with_exit_paragraph_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. FLOW-A.\nPROCEDURE DIVISION.\n    PERFORM STEP-A THRU STEP-C.\n    STOP RUN.\nSTEP-A.\n    DISPLAY \"A\".\nSTEP-B.\n    DISPLAY \"B\".\nSTEP-C.\n    DISPLAY \"C\".",
    );
}

#[test]
fn goto_and_recovery_label_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. FLOW-B.\nPROCEDURE DIVISION.\n    GO TO RECOVER-LABEL.\nMAIN-LABEL.\n    DISPLAY \"MAIN\".\n    STOP RUN.\nRECOVER-LABEL.\n    DISPLAY \"RECOVER\".\n    GO TO MAIN-LABEL.",
    );
}

#[test]
fn perform_until_with_inner_if_compiles() {
    compile_ok(&p(
        "01 WS-I PIC 9 VALUE 0.",
        "    PERFORM UNTIL WS-I >= 5\n        ADD 1 TO WS-I\n        IF WS-I = 3\n            DISPLAY \"MID\"\n        END-IF\n    END-PERFORM.",
    ));
}
