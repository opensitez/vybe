use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test] fn move_literal_to_field_compiles() { compile_ok(&p("01 WS-NAME PIC X(10).", "    MOVE \"ALICE\" TO WS-NAME.")); }
#[test] fn move_numeric_to_field_compiles() { compile_ok(&p("01 WS-NUM PIC 9(3).", "    MOVE 42 TO WS-NUM.")); }
#[test] fn move_spaces_to_field_compiles() { compile_ok(&p("01 WS-NAME PIC X(10).", "    MOVE SPACES TO WS-NAME.")); }
#[test] fn move_zeros_to_field_compiles() { compile_ok(&p("01 WS-NUM PIC 9(3).", "    MOVE ZEROS TO WS-NUM.")); }
#[test] fn move_field_to_field_compiles() { compile_ok(&p("01 WS-A PIC X(5) VALUE \"HI\".\n01 WS-B PIC X(5).", "    MOVE WS-A TO WS-B.")); }
#[test] fn add_to_statement_compiles() { compile_ok(&p("01 WS-A PIC 9(3) VALUE 2.\n01 WS-B PIC 9(3) VALUE 3.", "    ADD WS-A TO WS-B.")); }
#[test] fn add_giving_statement_compiles() { compile_ok(&p("01 WS-A PIC 9(3) VALUE 2.\n01 WS-B PIC 9(3) VALUE 3.\n01 WS-C PIC 9(3).", "    ADD WS-A WS-B GIVING WS-C.")); }
#[test] fn subtract_from_statement_compiles() { compile_ok(&p("01 WS-A PIC 9(3) VALUE 8.\n01 WS-B PIC 9(3) VALUE 3.", "    SUBTRACT WS-B FROM WS-A.")); }
#[test] fn subtract_giving_statement_compiles() { compile_ok(&p("01 WS-A PIC 9(3) VALUE 8.\n01 WS-B PIC 9(3) VALUE 3.\n01 WS-C PIC 9(3).", "    SUBTRACT WS-B FROM WS-A GIVING WS-C.")); }
#[test] fn multiply_by_statement_compiles() { compile_ok(&p("01 WS-A PIC 9(3) VALUE 4.\n01 WS-B PIC 9(3) VALUE 5.", "    MULTIPLY WS-A BY WS-B.")); }
#[test] fn multiply_giving_statement_compiles() { compile_ok(&p("01 WS-A PIC 9(3) VALUE 4.\n01 WS-B PIC 9(3) VALUE 5.\n01 WS-C PIC 9(3).", "    MULTIPLY WS-A BY WS-B GIVING WS-C.")); }
#[test] fn divide_by_statement_compiles() { compile_ok(&p("01 WS-A PIC 9(3) VALUE 9.\n01 WS-B PIC 9(3) VALUE 3.", "    DIVIDE WS-A BY WS-B.")); }
#[test] fn divide_giving_statement_compiles() { compile_ok(&p("01 WS-A PIC 9(3) VALUE 9.\n01 WS-B PIC 9(3) VALUE 3.\n01 WS-C PIC 9(3).", "    DIVIDE WS-A BY WS-B GIVING WS-C.")); }
#[test] fn compute_arithmetic_expression_compiles() { compile_ok(&p("01 WS-A PIC 9(3) VALUE 2.\n01 WS-B PIC 9(3) VALUE 3.\n01 WS-C PIC 9(3).", "    COMPUTE WS-C = WS-A + WS-B.")); }
#[test] fn compute_with_parentheses_compiles() { compile_ok(&p("01 WS-A PIC 9(3) VALUE 2.\n01 WS-B PIC 9(3) VALUE 3.\n01 WS-C PIC 9(3).", "    COMPUTE WS-C = (WS-A + WS-B) * 2.")); }
#[test] fn compute_with_exponent_compiles() { compile_ok(&p("01 WS-A PIC 9(3) VALUE 2.\n01 WS-B PIC 9(3) VALUE 3.\n01 WS-C PIC 9(3).", "    COMPUTE WS-C = WS-A ** WS-B.")); }
#[test] fn initialize_data_item_compiles() { compile_ok(&p("01 WS-NAME PIC X(5) VALUE \"AB\".", "    INITIALIZE WS-NAME.")); }
#[test] fn initialize_group_item_compiles() { compile_ok(&p("01 WS-REC.\n   05 WS-A PIC X(3) VALUE \"AB\".\n   05 WS-B PIC 9(2) VALUE 10.", "    INITIALIZE WS-REC.")); }
#[test] fn set_true_condition_name_compiles() { compile_ok(&p("01 WS-FLAG PIC 9(1).\n   88 WS-ON VALUE 1.", "    SET WS-ON TO TRUE.")); }
#[test] fn set_false_condition_name_compiles() { compile_ok(&p("01 WS-FLAG PIC 9(1).\n   88 WS-OFF VALUE 0.", "    SET WS-OFF TO FALSE.")); }
#[test] fn call_statement_compiles() { compile_ok(&p("", "    CALL \"SUBPROG\".")); }
#[test] fn call_with_using_compiles() { compile_ok(&p("01 WS-A PIC X(5).", "    CALL \"SUBPROG\" USING WS-A.")); }
#[test] fn display_literal_compiles() { compile_ok(&p("", "    DISPLAY \"HELLO\".")); }
#[test] fn display_multiple_items_compiles() { compile_ok(&p("01 WS-A PIC X(5) VALUE \"A\".\n01 WS-B PIC 9(2) VALUE 10.", "    DISPLAY WS-A WS-B.")); }
#[test] fn accept_statement_compiles() { compile_ok(&p("01 WS-A PIC X(10).", "    ACCEPT WS-A.")); }
#[test] fn accept_date_statement_compiles() { compile_ok(&p("01 WS-DATE PIC X(8).", "    ACCEPT WS-DATE FROM DATE.")); }
