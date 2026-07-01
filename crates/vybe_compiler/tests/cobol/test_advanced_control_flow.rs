use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test] fn evaluate_when_other_compiles() { compile_ok(&p("01 WS-A PIC 9(1) VALUE 9.", "    EVALUATE WS-A\n        WHEN 1\n            DISPLAY \"ONE\"\n        WHEN 2\n            DISPLAY \"TWO\"\n        WHEN OTHER\n            DISPLAY \"OTHER\"\n    END-EVALUATE.")); }
#[test] fn evaluate_multiple_branches_compiles() { compile_ok(&p("01 WS-A PIC 9(1) VALUE 2.", "    EVALUATE WS-A\n        WHEN 1\n            DISPLAY \"ONE\"\n        WHEN 2\n            DISPLAY \"TWO\"\n        WHEN 3\n            DISPLAY \"THREE\"\n    END-EVALUATE.")); }
#[test] fn evaluate_true_condition_compiles() { compile_ok(&p("01 WS-A PIC 9(2) VALUE 85.", "    EVALUATE TRUE\n        WHEN WS-A >= 90\n            DISPLAY \"A\"\n        WHEN WS-A >= 80\n            DISPLAY \"B\"\n        WHEN OTHER\n            DISPLAY \"F\"\n    END-EVALUATE.")); }
#[test]
fn perform_through_paragraphs_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    PERFORM PARA-ONE THRU PARA-THREE.\n    STOP RUN.\nPARA-ONE.\n    DISPLAY \"ONE\".\nPARA-TWO.\n    DISPLAY \"TWO\".\nPARA-THREE.\n    DISPLAY \"THREE\".",
    );
}

#[test]
fn goto_statement_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    GO TO LABEL-ONE.\nLABEL-ONE.\n    DISPLAY \"DONE\".\n    STOP RUN.",
    );
}

#[test]
fn alter_goto_target_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    ALTER LABEL-ONE TO PROCEED TO LABEL-TWO.\n    GO TO LABEL-ONE.\nLABEL-ONE.\n    DISPLAY \"ONE\".\nLABEL-TWO.\n    DISPLAY \"TWO\".\n    STOP RUN.",
    );
}
