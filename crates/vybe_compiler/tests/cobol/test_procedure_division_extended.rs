use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn procedure_division_add_statement_compiles() {
    compile_ok(&p("01 WS-A PIC 9(3) VALUE 5.", "    ADD 2 TO WS-A."));
}

#[test]
fn procedure_division_subtract_statement_compiles() {
    compile_ok(&p("01 WS-A PIC 9(3) VALUE 5.", "    SUBTRACT 2 FROM WS-A."));
}

#[test]
fn procedure_division_multiply_statement_compiles() {
    compile_ok(&p("01 WS-A PIC 9(3) VALUE 5.", "    MULTIPLY 2 BY WS-A."));
}

#[test]
fn procedure_division_divide_statement_compiles() {
    compile_ok(&p("01 WS-A PIC 9(3) VALUE 6.", "    DIVIDE 2 INTO WS-A."));
}

#[test]
fn procedure_division_display_statement_compiles() {
    compile_ok(&p("01 WS-A PIC X(3) VALUE \"ABC\".", "    DISPLAY WS-A."));
}

#[test]
fn procedure_division_move_statement_compiles() {
    compile_ok(&p("01 WS-A PIC X(3) VALUE \"ABC\".\n01 WS-B PIC X(3).", "    MOVE WS-A TO WS-B."));
}

#[test]
fn procedure_division_if_statement_compiles() {
    compile_ok(&p("01 WS-A PIC 9(3) VALUE 3.", "    IF WS-A > 0 DISPLAY \"POS\" END-IF."));
}

#[test]
fn procedure_division_perform_statement_compiles() {
    compile_ok(&p("", "    PERFORM 2 TIMES\n        DISPLAY \"X\"\n    END-PERFORM."));
}
