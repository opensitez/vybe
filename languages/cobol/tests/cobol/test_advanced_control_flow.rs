use super::helpers::run_prints;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn evaluate_when_other_compiles() {
    let out = run_prints(&p(
        "01 WS-A PIC 9(1) VALUE 9.",
        "    EVALUATE WS-A\n        WHEN 1\n            DISPLAY \"ONE\"\n        WHEN 2\n            DISPLAY \"TWO\"\n        WHEN OTHER\n            DISPLAY \"OTHER\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["OTHER"]);
}
#[test]
fn evaluate_multiple_branches_compiles() {
    let out = run_prints(&p(
        "01 WS-A PIC 9(1) VALUE 2.",
        "    EVALUATE WS-A\n        WHEN 1\n            DISPLAY \"ONE\"\n        WHEN 2\n            DISPLAY \"TWO\"\n        WHEN 3\n            DISPLAY \"THREE\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["TWO"]);
}
#[test]
fn evaluate_true_condition_compiles() {
    let out = run_prints(&p(
        "01 WS-A PIC 9(2) VALUE 85.",
        "    EVALUATE TRUE\n        WHEN WS-A >= 90\n            DISPLAY \"A\"\n        WHEN WS-A >= 80\n            DISPLAY \"B\"\n        WHEN OTHER\n            DISPLAY \"F\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["B"]);
}
#[test]
fn perform_through_paragraphs_compiles() {
    let out = run_prints(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    PERFORM PARA-ONE THRU PARA-THREE.\n    STOP RUN.\nPARA-ONE.\n    DISPLAY \"ONE\".\nPARA-TWO.\n    DISPLAY \"TWO\".\nPARA-THREE.\n    DISPLAY \"THREE\".",
    );
    assert_eq!(out, vec!["ONE", "TWO", "THREE"]);
}

#[test]
fn goto_statement_compiles() {
    let out = run_prints(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    GO TO LABEL-ONE.\nLABEL-ONE.\n    DISPLAY \"DONE\".\n    STOP RUN.",
    );
    assert_eq!(out, vec!["DONE"]);
}

#[test]
fn alter_goto_target_compiles() {
    let out = run_prints(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    ALTER LABEL-ONE TO PROCEED TO LABEL-TWO.\n    GO TO LABEL-ONE.\nLABEL-ONE.\n    DISPLAY \"ONE\".\nLABEL-TWO.\n    DISPLAY \"TWO\".\n    STOP RUN.",
    );
    assert_eq!(out, vec!["TWO"]);
}

#[test]
fn perform_twice_times_compiles() {
    let out = run_prints(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 COUNT PIC 9 VALUE 0.\nPROCEDURE DIVISION.\n    PERFORM 2 TIMES\n        ADD 1 TO COUNT\n        DISPLAY COUNT\n    END-PERFORM.\n    STOP RUN.",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn perform_with_test_after_compiles() {
    let out = run_prints(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 COUNT PIC 9 VALUE 0.\nPROCEDURE DIVISION.\n    PERFORM WITH TEST AFTER UNTIL COUNT > 1\n        ADD 1 TO COUNT\n        DISPLAY COUNT\n    END-PERFORM.\n    STOP RUN.",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn evaluate_with_elsif_path_compiles() {
    let out = run_prints(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-A PIC 9(1) VALUE 8.\nPROCEDURE DIVISION.\n    EVALUATE WS-A\n        WHEN 1 THRU 3\n            DISPLAY \"LOW\"\n        WHEN 4 THRU 7\n            DISPLAY \"MID\"\n        WHEN 8 THRU 9\n            DISPLAY \"HIGH\"\n        WHEN OTHER\n            DISPLAY \"OTHER\"\n    END-EVALUATE.\n    STOP RUN.",
    );
    assert_eq!(out, vec!["HIGH"]);
}

#[test]
fn goto_depending_compiles() {
    let out = run_prints(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 V PIC 9 VALUE 2.\nPROCEDURE DIVISION.\n    GO TO LABEL-ONE LABEL-TWO LABEL-THREE DEPENDING ON V.\n    DISPLAY \"OTHER\".\n    STOP RUN.\nLABEL-ONE.\n    DISPLAY \"ONE\".\n    STOP RUN.\nLABEL-TWO.\n    DISPLAY \"TWO\".\n    STOP RUN.\nLABEL-THREE.\n    DISPLAY \"THREE\".\n    STOP RUN.",
    );
    assert_eq!(out, vec!["TWO"]);
}
