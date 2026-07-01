use super::helpers::{compile_ok, compile_ok_check, parse_ok};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn call_basic() {
    compile_ok(&p(
        "",
        "    CALL \"SUBA\".",
    ));
}

#[test]
fn call_with_args() {
    compile_ok(&p(
        "01 WS-INPUT PIC X(20) VALUE \"Data\".",
        "    CALL \"SUBB\" USING WS-INPUT.",
    ));
}

#[test]
fn call_on_exception() {
    compile_ok(&p(
        "",
        "    CALL \"SUBC\"\n        ON EXCEPTION DISPLAY \"E\"\n        NOT ON EXCEPTION DISPLAY \"O\"\n    END-CALL.",
    ));
}

#[test]
fn perform_times_with_call() {
    compile_ok(&p("", "    PERFORM 2 TIMES\n        CALL \"SUBD\"\n    END-PERFORM."));
}

#[test]
fn perform_until_with_counter() {
    compile_ok(&p("01 I PIC 9 VALUE 0.", "    PERFORM UNTIL I >= 3\n        ADD 1 TO I\n        CALL \"SUBE\"\n    END-PERFORM."));
}

#[test]
fn evaluate_dispatch_calls() {
    compile_ok(&p(
        "01 K PIC 9 VALUE 1.",
        "    EVALUATE K\n        WHEN 1 CALL \"S1\"\n        WHEN 2 CALL \"S2\"\n        WHEN OTHER CALL \"SX\"\n    END-EVALUATE.",
    ));
}

#[test]
fn if_else_call() {
    compile_ok(&p(
        "01 F PIC 9 VALUE 1.",
        "    IF F = 1\n        CALL \"YESP\"\n    ELSE\n        CALL \"NOP\"\n    END-IF.",
    ));
}
