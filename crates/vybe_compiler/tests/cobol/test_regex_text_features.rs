use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test] fn inspect_count_letters_compiles() { compile_ok(&p("01 S PIC X(30).\n01 C PIC 9(3) VALUE 0.", "    INSPECT S TALLYING C FOR ALL \"A\".")); }
#[test] fn inspect_replace_letters_compiles() { compile_ok(&p("01 S PIC X(30).", "    INSPECT S REPLACING ALL \"A\" BY \"B\".")); }
#[test] fn inspect_first_replacement_compiles() { compile_ok(&p("01 S PIC X(30).", "    INSPECT S REPLACING FIRST \"A\" BY \"Z\".")); }
#[test] fn inspect_leading_count_compiles() { compile_ok(&p("01 S PIC X(30).\n01 C PIC 9(3) VALUE 0.", "    INSPECT S TALLYING C FOR LEADING \"0\".")); }
#[test] fn unstring_split_compiles() { compile_ok(&p("01 S PIC X(30).\n01 A PIC X(10).\n01 B PIC X(10).", "    UNSTRING S DELIMITED BY \",\" INTO A B.")); }
#[test] fn string_join_parts_compiles() { compile_ok(&p("01 A PIC X(10).\n01 B PIC X(10).\n01 O PIC X(20).", "    STRING A DELIMITED BY SIZE B DELIMITED BY SIZE INTO O.")); }
#[test] fn reference_modification_compiles() { compile_ok(&p("01 S PIC X(30).\n01 O PIC X(10).", "    MOVE S(1:10) TO O.")); }
#[test] fn text_trim_before_regex_compiles() { compile_ok(&p("01 S PIC X(30).\n01 O PIC X(30).", "    MOVE FUNCTION TRIM(S) TO O.")); }
#[test] fn text_upper_before_regex_compiles() { compile_ok(&p("01 S PIC X(30).\n01 O PIC X(30).", "    MOVE FUNCTION UPPER-CASE(S) TO O.")); }
#[test] fn text_lower_before_regex_compiles() { compile_ok(&p("01 S PIC X(30).\n01 O PIC X(30).", "    MOVE FUNCTION LOWER-CASE(S) TO O.")); }
#[test] fn text_length_for_regex_compiles() { compile_ok(&p("01 S PIC X(30).\n01 L PIC 9(3).", "    MOVE FUNCTION LENGTH(S) TO L.")); }
#[test] fn unstring_parts_compiles() { compile_ok(&p("01 S PIC X(20).\n01 A PIC X(10).\n01 B PIC X(10).", "    UNSTRING S DELIMITED BY \",\" INTO A B.")); }
