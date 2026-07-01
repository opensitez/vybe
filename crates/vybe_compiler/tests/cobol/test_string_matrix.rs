use super::helpers::compile_ok_check;

fn make_string_program(src: &str, target: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-SRC PIC X(20) VALUE \"{src}\".\n01 WS-TGT PIC X(20).\nPROCEDURE DIVISION.\n    {target}\n    STOP RUN."
    )
}

#[test]
fn test_string_and_unstring_matrix() {
    let programs = [
        make_string_program("HELLO", "STRING WS-SRC DELIMITED BY SIZE INTO WS-TGT."),
        make_string_program("HELLO", "STRING WS-SRC DELIMITED BY SPACE INTO WS-TGT."),
        make_string_program("HELLO", "STRING \"!\" DELIMITED BY SIZE WS-SRC DELIMITED BY SIZE INTO WS-TGT."),
        make_string_program("A,B,C", "UNSTRING WS-SRC DELIMITED BY \",\" INTO WS-TGT."),
        make_string_program("A B C", "UNSTRING WS-SRC DELIMITED BY SPACE INTO WS-TGT."),
        make_string_program("A,,B", "UNSTRING WS-SRC DELIMITED BY ALL \",\" INTO WS-TGT."),
        make_string_program("abc", "MOVE FUNCTION UPPER-CASE(WS-SRC) TO WS-TGT."),
        make_string_program("ABC", "MOVE FUNCTION LOWER-CASE(WS-SRC) TO WS-TGT."),
        make_string_program("  abc  ", "MOVE FUNCTION TRIM(WS-SRC) TO WS-TGT."),
        make_string_program("abc", "MOVE FUNCTION REVERSE(WS-SRC) TO WS-TGT."),
        make_string_program("abc", "MOVE FUNCTION LENGTH(WS-SRC) TO WS-TGT."),
    ];

    for program in programs {
        assert!(compile_ok_check(&program), "string case failed for program:\n{program}");
    }
}
