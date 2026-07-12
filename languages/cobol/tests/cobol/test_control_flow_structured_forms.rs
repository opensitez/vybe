use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn if_simple_compiles() {
    compile_ok(&p(
        "01 X PIC 9 VALUE 1.",
        "    IF X = 1 DISPLAY \"A\" END-IF.",
    ));
}
#[test]
fn if_else_compiles() {
    compile_ok(&p(
        "01 X PIC 9 VALUE 2.",
        "    IF X = 1 DISPLAY \"A\" ELSE DISPLAY \"B\" END-IF.",
    ));
}
#[test]
fn if_nested_compiles() {
    compile_ok(&p(
        "01 X PIC 9 VALUE 2.",
        "    IF X > 0\n        IF X < 5 DISPLAY \"Y\" END-IF\n    END-IF.",
    ));
}
#[test]
fn if_and_compiles() {
    compile_ok(&p(
        "01 A PIC 9 VALUE 1.\n01 B PIC 9 VALUE 1.",
        "    IF A = 1 AND B = 1 DISPLAY \"Y\" END-IF.",
    ));
}
#[test]
fn if_or_compiles() {
    compile_ok(&p(
        "01 A PIC 9 VALUE 0.\n01 B PIC 9 VALUE 1.",
        "    IF A = 1 OR B = 1 DISPLAY \"Y\" END-IF.",
    ));
}
#[test]
fn if_not_compiles() {
    compile_ok(&p(
        "01 A PIC 9 VALUE 0.",
        "    IF NOT A = 1 DISPLAY \"Y\" END-IF.",
    ));
}
#[test]
fn evaluate_basic_compiles() {
    compile_ok(&p(
        "01 X PIC 9 VALUE 2.",
        "    EVALUATE X\n        WHEN 1 DISPLAY \"A\"\n        WHEN 2 DISPLAY \"B\"\n        WHEN OTHER DISPLAY \"C\"\n    END-EVALUATE.",
    ));
}
#[test]
fn evaluate_true_compiles() {
    compile_ok(&p(
        "01 N PIC 9(2) VALUE 80.",
        "    EVALUATE TRUE\n        WHEN N >= 90 DISPLAY \"A\"\n        WHEN N >= 80 DISPLAY \"B\"\n        WHEN OTHER DISPLAY \"C\"\n    END-EVALUATE.",
    ));
}
#[test]
fn perform_times_compiles() {
    compile_ok(&p("", "    PERFORM 3 TIMES DISPLAY \"L\" END-PERFORM."));
}
#[test]
fn perform_until_compiles() {
    compile_ok(&p(
        "01 I PIC 9 VALUE 0.",
        "    PERFORM UNTIL I >= 3\n        ADD 1 TO I\n    END-PERFORM.",
    ));
}
#[test]
fn perform_varying_compiles() {
    compile_ok(&p(
        "01 I PIC 9 VALUE 0.",
        "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 3\n        DISPLAY I\n    END-PERFORM.",
    ));
}
#[test]
fn perform_paragraph_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    PERFORM P1.\n    STOP RUN.\nP1.\n    DISPLAY \"P\".",
    );
}
#[test]
fn perform_thru_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    PERFORM P1 THRU P2.\n    STOP RUN.\nP1.\n    DISPLAY \"1\".\nP2.\n    DISPLAY \"2\".",
    );
}
#[test]
fn continue_stmt_compiles() {
    compile_ok(&p("01 X PIC 9 VALUE 1.", "    IF X = 1 CONTINUE END-IF."));
}
#[test]
fn goto_stmt_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    GO TO L1.\nL1.\n    DISPLAY \"OK\".\n    STOP RUN.",
    );
}
#[test]
fn alter_stmt_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    ALTER L1 TO PROCEED TO L2.\n    GO TO L1.\nL1. DISPLAY \"A\".\nL2. DISPLAY \"B\".\n    STOP RUN.",
    );
}
#[test]
fn goback_stmt_compiles() {
    compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    GOBACK.");
}
#[test]
fn exit_program_stmt_compiles() {
    compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    EXIT PROGRAM.");
}
