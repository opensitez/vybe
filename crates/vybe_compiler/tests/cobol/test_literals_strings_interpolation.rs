use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test] fn literal_numeric_move_compiles() { compile_ok(&p("01 A PIC 9(5).", "    MOVE 12345 TO A.")); }
#[test] fn literal_string_move_compiles() { compile_ok(&p("01 A PIC X(10).", "    MOVE \"HELLO\" TO A.")); }
#[test] fn literal_spaces_move_compiles() { compile_ok(&p("01 A PIC X(10).", "    MOVE SPACES TO A.")); }
#[test] fn literal_zeros_move_compiles() { compile_ok(&p("01 A PIC 9(5).", "    MOVE ZEROS TO A.")); }
#[test] fn display_literal_compiles() { compile_ok(&p("", "    DISPLAY \"LIT\".")); }
#[test] fn display_literal_concat_compiles() { compile_ok(&p("01 N PIC X(5) VALUE \"A\".", "    DISPLAY \"X\" N \"Y\".")); }
#[test] fn string_delimited_size_compiles() { compile_ok(&p("01 A PIC X(5) VALUE \"AB\".\n01 B PIC X(5) VALUE \"CD\".\n01 O PIC X(10).", "    STRING A DELIMITED BY SIZE B DELIMITED BY SIZE INTO O.")); }
#[test] fn unstring_delimited_comma_compiles() { compile_ok(&p("01 S PIC X(10) VALUE \"A,B\".\n01 A PIC X(5).\n01 B PIC X(5).", "    UNSTRING S DELIMITED BY \",\" INTO A B.")); }
#[test] fn inspect_replacing_literal_compiles() { compile_ok(&p("01 S PIC X(10) VALUE \"ABA\".", "    INSPECT S REPLACING ALL \"A\" BY \"Z\".")); }
#[test] fn inspect_tallying_literal_compiles() { compile_ok(&p("01 S PIC X(10) VALUE \"ABA\".\n01 C PIC 9(3) VALUE 0.", "    INSPECT S TALLYING C FOR ALL \"A\".")); }
#[test] fn lower_case_function_compiles() { compile_ok(&p("01 S PIC X(10) VALUE \"ABC\".\n01 O PIC X(10).", "    MOVE FUNCTION LOWER-CASE(S) TO O.")); }
#[test] fn upper_case_function_compiles() { compile_ok(&p("01 S PIC X(10) VALUE \"abc\".\n01 O PIC X(10).", "    MOVE FUNCTION UPPER-CASE(S) TO O.")); }
#[test] fn trim_function_compiles() { compile_ok(&p("01 S PIC X(10) VALUE \"  HI  \".\n01 O PIC X(10).", "    MOVE FUNCTION TRIM(S) TO O.")); }
#[test] fn reverse_function_compiles() { compile_ok(&p("01 S PIC X(10) VALUE \"ABC\".\n01 O PIC X(10).", "    MOVE FUNCTION REVERSE(S) TO O.")); }
#[test] fn length_function_compiles() { compile_ok(&p("01 S PIC X(10) VALUE \"ABC\".\n01 L PIC 9(3).", "    MOVE FUNCTION LENGTH(S) TO L.")); }
#[test] fn move_literal_to_group_compiles() { compile_ok(&p("01 G.\n   05 A PIC X(3).\n   05 B PIC X(3).", "    MOVE \"ABCDEF\" TO G.")); }
#[test] fn display_group_literal_mix_compiles() { compile_ok(&p("01 G PIC X(5) VALUE \"GROUP\".", "    DISPLAY \"VAL:\" G.")); }
#[test] fn literal_in_if_compare_compiles() { compile_ok(&p("01 A PIC X(5) VALUE \"YES\".", "    IF A = \"YES\" DISPLAY \"OK\" END-IF.")); }
