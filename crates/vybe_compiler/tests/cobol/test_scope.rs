use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn nested_if_scope_compiles() {
    compile_ok(&p(
        "01 WS-X PIC 9(2) VALUE 5.",
        "    IF WS-X > 0\n        IF WS-X < 10\n            DISPLAY \"in-range\"\n        END-IF\n    END-IF.",
    ));
}

#[test]
fn evaluate_scope_compiles() {
    compile_ok(&p(
        "01 WS-VAL PIC 9(1) VALUE 2.",
        "    EVALUATE WS-VAL\n        WHEN 1\n            DISPLAY \"one\"\n        WHEN 2\n            DISPLAY \"two\"\n    END-EVALUATE.",
    ));
}

#[test]
fn nested_evaluate_if_scope_compiles() {
    compile_ok(&p(
        "01 V PIC 9 VALUE 2.",
        "    EVALUATE V\n        WHEN 2\n            IF V > 0\n                DISPLAY \"POS\"\n            END-IF\n        WHEN OTHER\n            DISPLAY \"OTHER\"\n    END-EVALUATE.",
    ));
}

#[test]
fn perform_if_scope_compiles() {
    compile_ok(&p(
        "01 I PIC 9 VALUE 0.\n01 F PIC 9 VALUE 1.",
        "    PERFORM UNTIL I >= 2\n        ADD 1 TO I\n        IF F = 1 DISPLAY I END-IF\n    END-PERFORM.",
    ));
}
