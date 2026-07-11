use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn if_case_01() {
    compile_ok(&p(
        "01 A PIC 9 VALUE 1.",
        "    IF A = 1 DISPLAY \"Y\" END-IF.",
    ));
}
#[test]
fn if_case_02() {
    compile_ok(&p(
        "01 A PIC 9 VALUE 1.",
        "    IF A = 2 DISPLAY \"N\" ELSE DISPLAY \"Y\" END-IF.",
    ));
}
#[test]
fn if_case_03() {
    compile_ok(&p(
        "01 A PIC 9 VALUE 1.\n01 B PIC 9 VALUE 2.",
        "    IF A = 1 AND B = 2 DISPLAY \"Y\" END-IF.",
    ));
}
#[test]
fn if_case_04() {
    compile_ok(&p(
        "01 A PIC 9 VALUE 0.\n01 B PIC 9 VALUE 2.",
        "    IF A = 1 OR B = 2 DISPLAY \"Y\" END-IF.",
    ));
}
#[test]
fn if_case_05() {
    compile_ok(&p(
        "01 A PIC 9 VALUE 0.",
        "    IF NOT A = 1 DISPLAY \"Y\" END-IF.",
    ));
}
#[test]
fn if_case_06() {
    compile_ok(&p(
        "01 X PIC X(3) VALUE \"123\".",
        "    IF X IS NUMERIC DISPLAY \"N\" END-IF.",
    ));
}
#[test]
fn if_case_07() {
    compile_ok(&p(
        "01 X PIC X(3) VALUE \"ABC\".",
        "    IF X IS ALPHABETIC DISPLAY \"A\" END-IF.",
    ));
}
#[test]
fn if_case_08() {
    compile_ok(&p(
        "01 X PIC S9 VALUE -1.",
        "    IF X IS NEGATIVE DISPLAY \"N\" END-IF.",
    ));
}
#[test]
fn if_case_09() {
    compile_ok(&p(
        "01 X PIC S9 VALUE 0.",
        "    IF X IS ZERO DISPLAY \"Z\" END-IF.",
    ));
}
#[test]
fn if_case_10() {
    compile_ok(&p(
        "01 X PIC S9 VALUE 3.",
        "    IF X IS POSITIVE DISPLAY \"P\" END-IF.",
    ));
}

#[test]
fn eval_case_01() {
    compile_ok(&p(
        "01 X PIC 9 VALUE 1.",
        "    EVALUATE X WHEN 1 DISPLAY \"A\" WHEN OTHER DISPLAY \"Z\" END-EVALUATE.",
    ));
}
#[test]
fn eval_case_02() {
    compile_ok(&p(
        "01 X PIC 9 VALUE 2.",
        "    EVALUATE TRUE WHEN X = 1 DISPLAY \"A\" WHEN X = 2 DISPLAY \"B\" WHEN OTHER DISPLAY \"Z\" END-EVALUATE.",
    ));
}
#[test]
fn eval_case_03() {
    compile_ok(&p(
        "01 X PIC 9 VALUE 7.",
        "    EVALUATE X WHEN 1 THRU 5 DISPLAY \"L\" WHEN 6 THRU 9 DISPLAY \"H\" WHEN OTHER DISPLAY \"O\" END-EVALUATE.",
    ));
}
#[test]
fn eval_case_04() {
    compile_ok(&p(
        "01 A PIC 9 VALUE 1.\n01 B PIC 9 VALUE 2.",
        "    EVALUATE A ALSO B WHEN 1 ALSO 2 DISPLAY \"M\" WHEN OTHER DISPLAY \"N\" END-EVALUATE.",
    ));
}
#[test]
fn eval_case_05() {
    compile_ok(&p(
        "01 A PIC 9 VALUE 5.\n01 B PIC 9 VALUE 9.",
        "    EVALUATE A ALSO B WHEN 5 ALSO ANY DISPLAY \"H\" WHEN OTHER DISPLAY \"N\" END-EVALUATE.",
    ));
}
#[test]
fn eval_case_06() {
    compile_ok(&p(
        "01 X PIC 9 VALUE 3.",
        "    EVALUATE X WHEN 1 DISPLAY \"A\" WHEN 2 DISPLAY \"B\" WHEN 3 DISPLAY \"C\" WHEN OTHER DISPLAY \"Z\" END-EVALUATE.",
    ));
}
#[test]
fn eval_case_07() {
    compile_ok(&p(
        "01 X PIC X VALUE \"B\".",
        "    EVALUATE X WHEN \"A\" DISPLAY \"1\" WHEN \"B\" DISPLAY \"2\" WHEN OTHER DISPLAY \"3\" END-EVALUATE.",
    ));
}
#[test]
fn eval_case_08() {
    compile_ok(&p(
        "01 X PIC 9 VALUE 0.",
        "    EVALUATE TRUE WHEN X = 0 DISPLAY \"Z\" WHEN OTHER DISPLAY \"N\" END-EVALUATE.",
    ));
}

#[test]
fn perf_case_01() {
    compile_ok(&p(
        "01 I PIC 9 VALUE 0.",
        "    PERFORM UNTIL I > 2 ADD 1 TO I END-PERFORM.",
    ));
}
#[test]
fn perf_case_02() {
    compile_ok(&p(
        "01 I PIC 9 VALUE 0.",
        "    PERFORM WITH TEST BEFORE UNTIL I > 2 ADD 1 TO I END-PERFORM.",
    ));
}
#[test]
fn perf_case_03() {
    compile_ok(&p(
        "01 I PIC 9 VALUE 0.",
        "    PERFORM WITH TEST AFTER UNTIL I > 2 ADD 1 TO I END-PERFORM.",
    ));
}
#[test]
fn perf_case_04() {
    compile_ok(&p(
        "01 I PIC 9 VALUE 0.",
        "    PERFORM 3 TIMES ADD 1 TO I END-PERFORM.",
    ));
}
#[test]
fn perf_case_05() {
    compile_ok(&p(
        "01 I PIC 9 VALUE 0.",
        "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 3 DISPLAY I END-PERFORM.",
    ));
}
#[test]
fn perf_case_06() {
    compile_ok(&p(
        "01 I PIC 9 VALUE 0.",
        "    PERFORM VARYING I FROM 3 BY -1 UNTIL I < 1 DISPLAY I END-PERFORM.",
    ));
}
#[test]
fn perf_case_07() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    PERFORM P1.\n    STOP RUN.\nP1. DISPLAY \"P\".",
    );
}
#[test]
fn perf_case_08() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    PERFORM P1 THRU P2.\n    STOP RUN.\nP1. DISPLAY \"1\".\nP2. DISPLAY \"2\".",
    );
}
#[test]
fn perf_case_09() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    PERFORM S1.\n    STOP RUN.\nS1 SECTION.\nP1. DISPLAY \"S\".",
    );
}
#[test]
fn perf_case_10() {
    compile_ok(&p(
        "01 I PIC 9 VALUE 0.",
        "    PERFORM 2 TIMES IF I = 0 DISPLAY \"A\" END-IF ADD 1 TO I END-PERFORM.",
    ));
}

#[test]
fn call_case_01() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CALL \"SUB1\".\n    STOP RUN.",
    );
}
#[test]
fn call_case_02() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 X PIC 9 VALUE 1.\nPROCEDURE DIVISION.\n    CALL \"SUB2\" USING X.\n    STOP RUN.",
    );
}
#[test]
fn call_case_03() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 X PIC 9 VALUE 1.\nPROCEDURE DIVISION.\n    CALL \"SUB3\" USING BY REFERENCE X.\n    STOP RUN.",
    );
}
#[test]
fn call_case_04() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 X PIC 9 VALUE 1.\nPROCEDURE DIVISION.\n    CALL \"SUB4\" USING BY CONTENT X.\n    STOP RUN.",
    );
}
#[test]
fn call_case_05() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 X PIC 9 VALUE 1.\nPROCEDURE DIVISION.\n    CALL \"SUB5\" USING BY VALUE X.\n    STOP RUN.",
    );
}
#[test]
fn call_case_06() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 R PIC 9(2).\nPROCEDURE DIVISION.\n    CALL \"SUB6\" RETURNING R.\n    STOP RUN.",
    );
}
#[test]
fn call_case_07() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 PGM PIC X(8) VALUE \"SUB7\".\nPROCEDURE DIVISION.\n    CALL PGM.\n    STOP RUN.",
    );
}
#[test]
fn call_case_08() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CALL \"SUB8\" ON EXCEPTION DISPLAY \"E\" NOT ON EXCEPTION DISPLAY \"O\" END-CALL.\n    STOP RUN.",
    );
}

#[test]
fn exit_case_01() {
    compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    STOP RUN.");
}
#[test]
fn exit_case_02() {
    compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    GOBACK.");
}
#[test]
fn exit_case_03() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CANCEL \"SUB9\".\n    STOP RUN.",
    );
}
#[test]
fn exit_case_04() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 PGM PIC X(8) VALUE \"SUB9\".\nPROCEDURE DIVISION.\n    CANCEL PGM.\n    STOP RUN.",
    );
}
#[test]
fn exit_case_05() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    GO TO L1.\nL1. STOP RUN.",
    );
}
#[test]
fn exit_case_06() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    ALTER L1 TO PROCEED TO L2.\nL1. DISPLAY \"A\".\nL2. STOP RUN.",
    );
}
#[test]
fn exit_case_07() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CONTINUE.\n    STOP RUN.",
    );
}
#[test]
fn exit_case_08() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    IF 1 = 1 CONTINUE END-IF.\n    STOP RUN.",
    );
}
