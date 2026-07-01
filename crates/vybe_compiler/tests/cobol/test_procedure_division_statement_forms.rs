use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test] fn move_stmt_compiles() { compile_ok(&p("01 A PIC X(5).", "    MOVE \"A\" TO A.")); }
#[test] fn add_stmt_compiles() { compile_ok(&p("01 A PIC 9 VALUE 1.\n01 B PIC 9 VALUE 1.", "    ADD A TO B.")); }
#[test] fn subtract_stmt_compiles() { compile_ok(&p("01 A PIC 9 VALUE 2.\n01 B PIC 9 VALUE 1.", "    SUBTRACT B FROM A.")); }
#[test] fn multiply_stmt_compiles() { compile_ok(&p("01 A PIC 9 VALUE 2.\n01 B PIC 9 VALUE 3.", "    MULTIPLY A BY B.")); }
#[test] fn divide_stmt_compiles() { compile_ok(&p("01 A PIC 9 VALUE 8.\n01 B PIC 9 VALUE 2.", "    DIVIDE A BY B.")); }
#[test] fn compute_stmt_compiles() { compile_ok(&p("01 A PIC 9 VALUE 2.\n01 B PIC 9 VALUE 3.\n01 C PIC 9.", "    COMPUTE C = A + B.")); }
#[test] fn initialize_stmt_compiles() { compile_ok(&p("01 G.\n   05 A PIC X(3) VALUE \"A\".", "    INITIALIZE G.")); }
#[test] fn set_true_stmt_compiles() { compile_ok(&p("01 F PIC 9.\n   88 ONN VALUE 1.", "    SET ONN TO TRUE.")); }
#[test] fn call_stmt_compiles() { compile_ok(&p("", "    CALL \"W\".")); }
#[test] fn display_stmt_compiles() { compile_ok(&p("01 A PIC X(5) VALUE \"A\".", "    DISPLAY A.")); }
#[test] fn accept_stmt_compiles() { compile_ok(&p("01 A PIC X(10).", "    ACCEPT A.")); }
#[test] fn perform_stmt_compiles() { compile_ok(&p("", "    PERFORM 2 TIMES DISPLAY \"L\" END-PERFORM.")); }
#[test] fn evaluate_stmt_compiles() { compile_ok(&p("01 X PIC 9 VALUE 1.", "    EVALUATE X WHEN 1 DISPLAY \"A\" WHEN OTHER DISPLAY \"B\" END-EVALUATE.")); }
#[test] fn if_stmt_compiles() { compile_ok(&p("01 X PIC 9 VALUE 1.", "    IF X = 1 DISPLAY \"A\" END-IF.")); }
#[test] fn goto_stmt_compiles() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    GO TO L1.\nL1.\n    STOP RUN."); }
#[test] fn alter_stmt_compiles() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    ALTER L1 TO PROCEED TO L2.\nL1. DISPLAY \"A\".\nL2. STOP RUN."); }
#[test] fn continue_stmt_compiles() { compile_ok(&p("01 X PIC 9 VALUE 1.", "    IF X = 1 CONTINUE END-IF.")); }
#[test] fn goback_stmt_compiles() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    GOBACK."); }
