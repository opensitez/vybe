use vybec::parser_cobol::parse;
use vybec::compiler_cobol::Compiler;

fn compile_ok(src: &str) {
    let program = parse(src).expect("parse failed");
    let c = Compiler::new();
    let res = c.compile(&program);
    assert!(res.is_ok(), "compile failed for:\n{}\nerror: {:?}", src, res.err());
}

fn p(data: &str, body: &str) -> String {
    format!("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.", data, body)
}

// ── IF/ELSE/END-IF ─────────────────────────────────────────
#[test] fn if_simple() { compile_ok(&p("01 X PIC 9(3) VALUE 5.", "    IF X > 3\n        DISPLAY \"Yes\"\n    END-IF.")); }
#[test] fn if_else() { compile_ok(&p("01 X PIC 9(3) VALUE 5.", "    IF X > 10\n        DISPLAY \"Big\"\n    ELSE\n        DISPLAY \"Small\"\n    END-IF.")); }
#[test] fn if_nested() { compile_ok(&p("01 X PIC 9(3) VALUE 5.", "    IF X > 10\n        DISPLAY \"A\"\n    ELSE\n        IF X > 5\n            DISPLAY \"B\"\n        ELSE\n            DISPLAY \"C\"\n        END-IF\n    END-IF.")); }
#[test] fn if_equal() { compile_ok(&p("01 X PIC 9(3) VALUE 5.", "    IF X = 5\n        DISPLAY \"Five\"\n    END-IF.")); }
#[test] fn if_not_equal() { compile_ok(&p("01 X PIC 9(3) VALUE 5.", "    IF X NOT = 0\n        DISPLAY \"Non-zero\"\n    END-IF.")); }
#[test] fn if_and() { compile_ok(&p("01 X PIC 9(3) VALUE 5.\n01 Y PIC 9(3) VALUE 10.", "    IF X > 0 AND Y > 0\n        DISPLAY \"Both positive\"\n    END-IF.")); }
#[test] fn if_or() { compile_ok(&p("01 X PIC 9(3) VALUE 5.", "    IF X = 5 OR X = 10\n        DISPLAY \"Match\"\n    END-IF.")); }
#[test] fn if_not() { compile_ok(&p("01 X PIC 9(3) VALUE 0.", "    IF NOT X > 0\n        DISPLAY \"Zero or negative\"\n    END-IF.")); }
#[test] fn if_le() { compile_ok(&p("01 X PIC 9(3) VALUE 5.", "    IF X <= 10\n        DISPLAY \"OK\"\n    END-IF.")); }
#[test] fn if_ge() { compile_ok(&p("01 X PIC 9(3) VALUE 5.", "    IF X >= 1\n        DISPLAY \"OK\"\n    END-IF.")); }

// ── EVALUATE ───────────────────────────────────────────────
#[test] fn eval_simple() { compile_ok(&p("01 X PIC 9(1) VALUE 2.", "    EVALUATE X\n        WHEN 1\n            DISPLAY \"One\"\n        WHEN 2\n            DISPLAY \"Two\"\n        WHEN OTHER\n            DISPLAY \"Other\"\n    END-EVALUATE.")); }
#[test] fn eval_true() { compile_ok(&p("01 X PIC 9(3) VALUE 85.", "    EVALUATE TRUE\n        WHEN X >= 90\n            DISPLAY \"A\"\n        WHEN X >= 80\n            DISPLAY \"B\"\n        WHEN X >= 70\n            DISPLAY \"C\"\n        WHEN OTHER\n            DISPLAY \"F\"\n    END-EVALUATE.")); }
#[test] fn eval_string() { compile_ok(&p("01 X PIC X(5) VALUE \"B\".", "    EVALUATE X\n        WHEN \"A\"\n            DISPLAY \"Alpha\"\n        WHEN \"B\"\n            DISPLAY \"Beta\"\n    END-EVALUATE.")); }
#[test] fn eval_no_other() { compile_ok(&p("01 X PIC 9(1) VALUE 1.", "    EVALUATE X\n        WHEN 1\n            DISPLAY \"One\"\n        WHEN 2\n            DISPLAY \"Two\"\n    END-EVALUATE.")); }

// ── PERFORM TIMES ──────────────────────────────────────────
#[test] fn perform_1_time() { compile_ok(&p("", "    PERFORM 1 TIMES\n        DISPLAY \"Once\"\n    END-PERFORM.")); }
#[test] fn perform_10_times() { compile_ok(&p("", "    PERFORM 10 TIMES\n        DISPLAY \"Loop\"\n    END-PERFORM.")); }
#[test] fn perform_nested_times() { compile_ok(&p("", "    PERFORM 3 TIMES\n        PERFORM 2 TIMES\n            DISPLAY \"Inner\"\n        END-PERFORM\n    END-PERFORM.")); }

// ── PERFORM UNTIL ──────────────────────────────────────────
#[test] fn perform_until_simple() { compile_ok(&p("01 I PIC 9(3) VALUE 0.", "    PERFORM UNTIL I >= 10\n        ADD 1 TO I\n    END-PERFORM.")); }
#[test] fn perform_until_eq() { compile_ok(&p("01 I PIC 9(3) VALUE 0.", "    PERFORM UNTIL I = 5\n        ADD 1 TO I\n    END-PERFORM.")); }

// ── PERFORM VARYING ────────────────────────────────────────
#[test] fn perform_varying_basic() { compile_ok(&p("01 I PIC 9(3) VALUE 0.", "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 10\n        DISPLAY I\n    END-PERFORM.")); }
#[test] fn perform_varying_step2() { compile_ok(&p("01 I PIC 9(3) VALUE 0.", "    PERFORM VARYING I FROM 0 BY 2 UNTIL I > 20\n        DISPLAY I\n    END-PERFORM.")); }
#[test] fn perform_varying_down() { compile_ok(&p("01 I PIC 9(3) VALUE 0.", "    PERFORM VARYING I FROM 10 BY -1 UNTIL I < 1\n        DISPLAY I\n    END-PERFORM.")); }

// ── PERFORM PARAGRAPH ──────────────────────────────────────
#[test] fn perform_para() {
    compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    PERFORM MY-PARA.\n    STOP RUN.\nMY-PARA.\n    DISPLAY \"Hello\".");
}
#[test] fn perform_thru() {
    compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    PERFORM INIT-PARA THRU CLEANUP-PARA.\n    STOP RUN.\nINIT-PARA.\n    DISPLAY \"Init\".\nCLEANUP-PARA.\n    DISPLAY \"Cleanup\".");
}
#[test] fn perform_multiple_paras() {
    compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    PERFORM STEP1-PARA.\n    PERFORM STEP2-PARA.\n    PERFORM STEP3-PARA.\n    STOP RUN.\nSTEP1-PARA.\n    DISPLAY \"Step 1\".\nSTEP2-PARA.\n    DISPLAY \"Step 2\".\nSTEP3-PARA.\n    DISPLAY \"Step 3\".");
}

// ── CONTINUE ───────────────────────────────────────────────
#[test] fn continue_in_if() { compile_ok(&p("01 X PIC 9(3) VALUE 5.", "    IF X > 10\n        DISPLAY \"Big\"\n    ELSE\n        CONTINUE\n    END-IF.")); }

// ── GOBACK ─────────────────────────────────────────────────
#[test] fn goback_stmt() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    DISPLAY \"Done\".\n    GOBACK."); }

// ── GO TO ──────────────────────────────────────────────────
#[test] fn goto_para() {
    compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    DISPLAY \"Start\".\n    STOP RUN.\nERROR-PARA.\n    DISPLAY \"Error\".");
}
