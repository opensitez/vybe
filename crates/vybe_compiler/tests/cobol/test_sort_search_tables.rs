use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test] fn sort_ascending_key_compiles() { compile_ok(&p("01 F PIC X(10).\n01 K PIC 9(5).", "    SORT F ON ASCENDING KEY K.")); }
#[test] fn sort_descending_key_compiles() { compile_ok(&p("01 F PIC X(10).\n01 K PIC 9(5).", "    SORT F ON DESCENDING KEY K.")); }
#[test] fn merge_ascending_key_compiles() { compile_ok(&p("01 F PIC X(10).\n01 K PIC 9(5).", "    MERGE F ON ASCENDING KEY K.")); }
#[test] fn merge_descending_key_compiles() { compile_ok(&p("01 F PIC X(10).\n01 K PIC 9(5).", "    MERGE F ON DESCENDING KEY K.")); }
#[test] fn search_basic_compiles() { compile_ok(&p("01 T.\n   05 E OCCURS 5 TIMES INDEXED BY I.\n      10 K PIC X(3).", "    SEARCH E WHEN K(I) = \"A\" DISPLAY \"F\" END-SEARCH.")); }
#[test] fn search_with_at_end_compiles() { compile_ok(&p("01 T.\n   05 E OCCURS 5 TIMES INDEXED BY I.\n      10 K PIC X(3).", "    SEARCH E AT END DISPLAY \"N\" WHEN K(I) = \"A\" DISPLAY \"F\" END-SEARCH.")); }
#[test] fn search_all_basic_compiles() { compile_ok(&p("01 T.\n   05 E OCCURS 5 TIMES ASCENDING KEY K INDEXED BY I.\n      10 K PIC X(3).", "    SEARCH ALL E WHEN K(I) = \"A\" DISPLAY \"F\" END-SEARCH.")); }
#[test] fn search_all_numeric_compiles() { compile_ok(&p("01 T.\n   05 E OCCURS 5 TIMES ASCENDING KEY K INDEXED BY I.\n      10 K PIC 9(3).", "    SEARCH ALL E WHEN K(I) = 100 DISPLAY \"F\" END-SEARCH.")); }
#[test] fn sort_then_search_pattern_compiles() { compile_ok(&p("01 F PIC X(10).\n01 K PIC 9(5).\n01 T.\n   05 E OCCURS 5 TIMES ASCENDING KEY K2 INDEXED BY I.\n      10 K2 PIC 9(5).", "    SORT F ON ASCENDING KEY K.\n    SEARCH ALL E WHEN K2(I) = 10 DISPLAY \"F\" END-SEARCH.")); }
#[test] fn search_loop_wrapper_compiles() { compile_ok(&p("01 N PIC 9 VALUE 0.\n01 T.\n   05 E OCCURS 5 TIMES INDEXED BY I.\n      10 K PIC X(3).", "    PERFORM UNTIL N >= 2\n        ADD 1 TO N\n        SEARCH E WHEN K(I) = \"A\" DISPLAY \"F\" END-SEARCH\n    END-PERFORM.")); }
#[test] fn sort_with_output_proc_compiles() { compile_ok(&p("01 F PIC X(10).\n01 K PIC 9(5).", "    SORT F ON ASCENDING KEY K OUTPUT PROCEDURE IS P1.")); }
#[test] fn sort_with_input_proc_compiles() { compile_ok(&p("01 F PIC X(10).\n01 K PIC 9(5).", "    SORT F ON ASCENDING KEY K INPUT PROCEDURE IS P1.")); }
#[test] fn merge_with_using_compiles() { compile_ok(&p("01 F PIC X(10).\n01 G PIC X(10).\n01 K PIC 9(5).", "    MERGE F ON ASCENDING KEY K USING G.")); }
#[test] fn search_all_with_if_compiles() { compile_ok(&p("01 T.\n   05 E OCCURS 5 TIMES ASCENDING KEY K INDEXED BY I.\n      10 K PIC 9(3).", "    IF 1 = 1 SEARCH ALL E WHEN K(I) = 1 DISPLAY \"F\" END-SEARCH END-IF.")); }
#[test] fn search_with_evaluate_compiles() { compile_ok(&p("01 T.\n   05 E OCCURS 5 TIMES INDEXED BY I.\n      10 K PIC X(3).", "    EVALUATE TRUE WHEN 1 = 1 SEARCH E WHEN K(I) = \"A\" DISPLAY \"F\" END-SEARCH END-EVALUATE.")); }
#[test] fn sort_key_group_compiles() { compile_ok(&p("01 F PIC X(10).\n01 R.\n   05 K PIC 9(5).", "    SORT F ON ASCENDING KEY K.")); }
#[test] fn search_key_group_compiles() { compile_ok(&p("01 T.\n   05 E OCCURS 5 TIMES INDEXED BY I.\n      10 R.\n         15 K PIC X(3).", "    SEARCH E WHEN K(I) = \"A\" DISPLAY \"F\" END-SEARCH.")); }
#[test] fn search_all_key_group_compiles() { compile_ok(&p("01 T.\n   05 E OCCURS 5 TIMES ASCENDING KEY K INDEXED BY I.\n      10 K PIC X(3).", "    SEARCH ALL E WHEN K(I) = \"A\" DISPLAY \"F\" END-SEARCH.")); }
