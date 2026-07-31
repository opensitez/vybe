use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn sort_ascending_key_compiles() {
    compile_ok(&p(
        "01 F PIC X(10).\n01 K PIC 9(5).",
        "    SORT F ON ASCENDING KEY K.",
    ));
}
#[test]
fn sort_descending_key_compiles() {
    compile_ok(&p(
        "01 F PIC X(10).\n01 K PIC 9(5).",
        "    SORT F ON DESCENDING KEY K.",
    ));
}
#[test]
fn merge_ascending_key_compiles() {
    compile_ok(&p(
        "01 F PIC X(10).\n01 K PIC 9(5).",
        "    MERGE F ON ASCENDING KEY K.",
    ));
}
#[test]
fn merge_descending_key_compiles() {
    compile_ok(&p(
        "01 F PIC X(10).\n01 K PIC 9(5).",
        "    MERGE F ON DESCENDING KEY K.",
    ));
}
#[test]
fn search_basic_compiles() {
    let out = run_prints(&p(
        "01 T.\n   05 E OCCURS 5 TIMES INDEXED BY I.\n      10 K PIC X(3).",
        "    MOVE \"A\" TO K(1).\n    MOVE \"B\" TO K(2).\n    MOVE \"C\" TO K(3).\n    SEARCH E WHEN K(I) = \"B\" DISPLAY \"FOUND\" END-SEARCH.",
    ));
    assert_eq!(out, vec!["FOUND"]);
}
#[test]
fn search_with_at_end_compiles() {
    let out = run_prints(&p(
        "01 T.\n   05 E OCCURS 5 TIMES INDEXED BY I.\n      10 K PIC X(3).",
        "    MOVE \"A\" TO K(1).\n    SEARCH E AT END DISPLAY \"N\" WHEN K(I) = \"Z\" DISPLAY \"FOUND\" END-SEARCH.",
    ));
    assert_eq!(out, vec!["N"]);
}
#[test]
fn search_all_basic_compiles() {
    let out = run_prints(&p(
        "01 T.\n   05 E OCCURS 5 TIMES ASCENDING KEY K INDEXED BY I.\n      10 K PIC X(3).",
        "    MOVE \"AA\" TO K(1).\n    MOVE \"BB\" TO K(2).\n    MOVE \"CC\" TO K(3).\n    SEARCH ALL E WHEN K(I) = \"BB\" DISPLAY \"FOUND\" END-SEARCH.",
    ));
    assert_eq!(out, vec!["FOUND"]);
}
#[test]
fn search_all_numeric_compiles() {
    let out = run_prints(&p(
        "01 T.\n   05 E OCCURS 5 TIMES ASCENDING KEY K INDEXED BY I.\n      10 K PIC 9(3).",
        "    MOVE 100 TO K(1).\n    MOVE 200 TO K(2).\n    MOVE 300 TO K(3).\n    SEARCH ALL E WHEN K(I) = 100 DISPLAY \"FOUND\" END-SEARCH.",
    ));
    assert_eq!(out, vec!["FOUND"]);
}
#[test]
fn sort_then_search_pattern_compiles() {
    compile_ok(&p(
        "01 F PIC X(10).\n01 K PIC 9(5).\n01 T.\n   05 E OCCURS 5 TIMES ASCENDING KEY K2 INDEXED BY I.\n      10 K2 PIC 9(5).",
        "    SORT F ON ASCENDING KEY K.\n    SEARCH ALL E WHEN K2(I) = 10 DISPLAY \"F\" END-SEARCH.",
    ));
}
#[test]
fn search_loop_wrapper_compiles() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 0.\n01 T.\n   05 E OCCURS 5 TIMES INDEXED BY I.\n      10 K PIC X(3).",
        "    MOVE \"A\" TO K(1).\n    MOVE \"A\" TO K(2).\n    SET I TO 1.\n    PERFORM UNTIL N >= 2\n        ADD 1 TO N\n        SEARCH E WHEN K(I) = \"A\" DISPLAY N END-SEARCH\n    END-PERFORM.",
    ));
    assert_eq!(out, vec!["1", "2"]);
}
#[test]
fn sort_with_output_proc_compiles() {
    compile_ok(&p(
        "01 F PIC X(10).\n01 K PIC 9(5).",
        "    SORT F ON ASCENDING KEY K OUTPUT PROCEDURE IS P1.",
    ));
}
#[test]
fn sort_with_input_proc_compiles() {
    compile_ok(&p(
        "01 F PIC X(10).\n01 K PIC 9(5).",
        "    SORT F ON ASCENDING KEY K INPUT PROCEDURE IS P1.",
    ));
}
#[test]
fn merge_with_using_compiles() {
    compile_ok(&p(
        "01 F PIC X(10).\n01 G PIC X(10).\n01 K PIC 9(5).",
        "    MERGE F ON ASCENDING KEY K USING G.",
    ));
}
#[test]
fn search_all_with_if_compiles() {
    let out = run_prints(&p(
        "01 T.\n   05 E OCCURS 5 TIMES ASCENDING KEY K INDEXED BY I.\n      10 K PIC 9(3).",
        "    MOVE 1 TO K(1).\n    MOVE 2 TO K(2).\n    IF 1 = 1\n        SEARCH ALL E WHEN K(I) = 1 DISPLAY \"YES\"\n    END-SEARCH\n    END-IF.",
    ));
    assert_eq!(out, vec!["YES"]);
}
#[test]
fn search_with_evaluate_compiles() {
    let out = run_prints(&p(
        "01 T.\n   05 E OCCURS 5 TIMES INDEXED BY I.\n      10 K PIC X(3).",
        "    MOVE \"A\" TO K(1).\n    MOVE \"B\" TO K(2).\n    EVALUATE TRUE\n    WHEN 1 = 1 SEARCH E WHEN K(I) = \"B\" DISPLAY \"EVAL\"\n    END-SEARCH\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["EVAL"]);
}
#[test]
fn sort_key_group_compiles() {
    compile_ok(&p(
        "01 F PIC X(10).\n01 R.\n   05 K PIC 9(5).",
        "    SORT F ON ASCENDING KEY K.",
    ));
}
#[test]
fn search_key_group_compiles() {
    let out = run_prints(&p(
        "01 T.\n   05 E OCCURS 5 TIMES INDEXED BY I.\n      10 R.\n         15 K PIC X(3).",
        "    MOVE \"A\" TO K(1).\n    MOVE \"B\" TO K(2).\n    SEARCH E WHEN K(I) = \"B\" DISPLAY \"FOUND\" END-SEARCH.",
    ));
    assert_eq!(out, vec!["FOUND"]);
}
#[test]
fn search_all_key_group_compiles() {
    let out = run_prints(&p(
        "01 T.\n   05 E OCCURS 5 TIMES ASCENDING KEY K INDEXED BY I.\n      10 K PIC X(3).",
        "    MOVE \"A\" TO K(1).\n    MOVE \"B\" TO K(2).\n    SEARCH ALL E WHEN K(I) = \"C\" DISPLAY \"FOUND\" END-SEARCH.\n    DISPLAY \"END\".",
    ));
    assert_eq!(out, vec!["END"]);
}

#[test]
fn search_not_found_runtime() {
    let out = run_prints(&p(
        "01 T.\n   05 E OCCURS 3 TIMES ASCENDING KEY K INDEXED BY I.\n      10 K PIC 9(3).",
        "    MOVE 1 TO K(1).\n    MOVE 2 TO K(2).\n    MOVE 3 TO K(3).\n    SEARCH ALL E WHEN K(I) = 9 DISPLAY 'FOUND' END-SEARCH.\n    DISPLAY 'END'.",
    ));
    assert_eq!(out, vec!["END"]);
}

#[test]
fn search_with_set_indexing_runtime() {
    let out = run_prints(&p(
        "01 T.\n   05 E OCCURS 4 TIMES INDEXED BY I.\n      10 K PIC 9(2).",
        "    MOVE 10 TO K(1).\n    MOVE 20 TO K(2).\n    MOVE 30 TO K(3).\n    SET I TO 2.\n    SEARCH E WHEN K(I) = 30 DISPLAY 'LATE' END-SEARCH.",
    ));
    assert_eq!(out, vec!["LATE"]);
}
