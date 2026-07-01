use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test] fn add_single_target_compiles() { compile_ok(&p("01 A PIC 9(3) VALUE 1.\n01 B PIC 9(3) VALUE 2.", "    ADD A TO B.")); }
#[test] fn add_giving_compiles() { compile_ok(&p("01 A PIC 9(3) VALUE 1.\n01 B PIC 9(3) VALUE 2.\n01 C PIC 9(3).", "    ADD A B GIVING C.")); }
#[test] fn subtract_single_compiles() { compile_ok(&p("01 A PIC 9(3) VALUE 9.\n01 B PIC 9(3) VALUE 2.", "    SUBTRACT B FROM A.")); }
#[test] fn subtract_giving_compiles() { compile_ok(&p("01 A PIC 9(3) VALUE 9.\n01 B PIC 9(3) VALUE 2.\n01 C PIC 9(3).", "    SUBTRACT B FROM A GIVING C.")); }
#[test] fn multiply_basic_compiles() { compile_ok(&p("01 A PIC 9(3) VALUE 3.\n01 B PIC 9(3) VALUE 2.", "    MULTIPLY A BY B.")); }
#[test] fn multiply_giving_compiles() { compile_ok(&p("01 A PIC 9(3) VALUE 3.\n01 B PIC 9(3) VALUE 2.\n01 C PIC 9(3).", "    MULTIPLY A BY B GIVING C.")); }
#[test] fn divide_basic_compiles() { compile_ok(&p("01 A PIC 9(3) VALUE 8.\n01 B PIC 9(3) VALUE 2.", "    DIVIDE A BY B.")); }
#[test] fn divide_giving_compiles() { compile_ok(&p("01 A PIC 9(3) VALUE 8.\n01 B PIC 9(3) VALUE 2.\n01 C PIC 9(3).", "    DIVIDE A BY B GIVING C.")); }
#[test] fn compute_plus_compiles() { compile_ok(&p("01 A PIC 9(3) VALUE 2.\n01 B PIC 9(3) VALUE 3.\n01 C PIC 9(3).", "    COMPUTE C = A + B.")); }
#[test] fn compute_minus_compiles() { compile_ok(&p("01 A PIC 9(3) VALUE 5.\n01 B PIC 9(3) VALUE 3.\n01 C PIC 9(3).", "    COMPUTE C = A - B.")); }
#[test] fn compute_times_compiles() { compile_ok(&p("01 A PIC 9(3) VALUE 5.\n01 B PIC 9(3) VALUE 3.\n01 C PIC 9(4).", "    COMPUTE C = A * B.")); }
#[test] fn compute_div_compiles() { compile_ok(&p("01 A PIC 9(3) VALUE 8.\n01 B PIC 9(3) VALUE 2.\n01 C PIC 9(4).", "    COMPUTE C = A / B.")); }
#[test] fn compute_parentheses_compiles() { compile_ok(&p("01 A PIC 9(3) VALUE 2.\n01 B PIC 9(3) VALUE 3.\n01 C PIC 9(4).", "    COMPUTE C = (A + B) * 2.")); }
#[test] fn compute_chain_compiles() { compile_ok(&p("01 A PIC 9(3) VALUE 2.\n01 B PIC 9(3) VALUE 3.\n01 C PIC 9(4).\n01 D PIC 9(5).", "    COMPUTE D = A + B + C.")); }
#[test] fn compute_exp_compiles() { compile_ok(&p("01 A PIC 9(3) VALUE 2.\n01 B PIC 9(3) VALUE 3.\n01 C PIC 9(5).", "    COMPUTE C = A ** B.")); }
#[test] fn numeric_move_zero_compiles() { compile_ok(&p("01 A PIC 9(5).", "    MOVE ZEROS TO A.")); }
#[test] fn numeric_move_literal_compiles() { compile_ok(&p("01 A PIC 9(5).", "    MOVE 12345 TO A.")); }
#[test] fn numeric_round_trip_compute_compiles() { compile_ok(&p("01 A PIC 9(3) VALUE 9.\n01 B PIC 9(3) VALUE 4.\n01 C PIC 9(3).", "    COMPUTE C = A - B.\n    ADD C TO B.")); }
