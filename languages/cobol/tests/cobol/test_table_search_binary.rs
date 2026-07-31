use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn search_all_ascending_key_compiles() {
    compile_ok(&p(
        "01 T.\n   05 E OCCURS 10 TIMES ASCENDING KEY E INDEXED BY IX.\n      10 E PIC 9(4).",
        "    SEARCH ALL E\n        AT END DISPLAY \"NOT FOUND\"\n        WHEN E(IX) = 5\n            DISPLAY \"FOUND\"\n    END-SEARCH.",
    ));
}

#[test]
fn search_all_string_key_compiles() {
    compile_ok(&p(
        "01 DICT.\n   05 ENTRY OCCURS 20 TIMES ASCENDING KEY WORD INDEXED BY DI.\n      10 WORD PIC X(10).\n      10 DEFN PIC X(40).",
        "    SEARCH ALL ENTRY\n        AT END DISPLAY \"MISSING\"\n        WHEN WORD(DI) = \"COBOL     \"\n            DISPLAY DEFN(DI)\n    END-SEARCH.",
    ));
}

#[test]
fn search_all_numeric_at_end() {
    compile_ok(&p(
        "01 NUMS.\n   05 NUM-ENTRY OCCURS 50 TIMES ASCENDING KEY NUM-VAL INDEXED BY NI.\n      10 NUM-VAL PIC 9(6).",
        "    SEARCH ALL NUM-ENTRY\n        AT END\n            DISPLAY \"VALUE NOT IN TABLE\"\n        WHEN NUM-VAL(NI) = 999999\n            DISPLAY \"FOUND MAX\"\n    END-SEARCH.",
    ));
}

#[test]
fn search_all_with_action_on_found() {
    compile_ok(&p(
        "01 PRODUCTS.\n   05 PROD OCCURS 100 TIMES ASCENDING KEY PROD-CODE INDEXED BY PI.\n      10 PROD-CODE PIC 9(5).\n      10 PROD-DESC PIC X(20).\n      10 PROD-PRICE PIC 9(7)V99.",
        "    SEARCH ALL PROD\n        AT END\n            DISPLAY \"NOT FOUND\"\n        WHEN PROD-CODE(PI) = 10001\n            DISPLAY PROD-DESC(PI)\n    END-SEARCH.",
    ));
}

#[test]
fn search_all_descending_key_compiles() {
    compile_ok(&p(
        "01 REVERSE-T.\n   05 RE OCCURS 10 TIMES DESCENDING KEY RE INDEXED BY RI.\n      10 RE PIC 9(4).",
        "    SEARCH ALL RE\n        AT END DISPLAY \"END\"\n        WHEN RE(RI) = 1\n            DISPLAY \"ONE\"\n    END-SEARCH.",
    ));
}

#[test]
fn search_all_compound_key_compiles() {
    compile_ok(&p(
        "01 COMPOUND-T.\n   05 CT OCCURS 20 TIMES ASCENDING KEY CT-KEY1 CT-KEY2 INDEXED BY CI.\n      10 CT-KEY1 PIC X(4).\n      10 CT-KEY2 PIC 9(4).\n      10 CT-DATA PIC X(20).",
        "    SEARCH ALL CT\n        AT END DISPLAY \"NOT FOUND\"\n        WHEN CT-KEY1(CI) = \"ABCD\" AND CT-KEY2(CI) = 1001\n            DISPLAY CT-DATA(CI)\n    END-SEARCH.",
    ));
}

#[test]
fn search_linear_found_at_first() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC X(3) OCCURS 5 TIMES INDEXED BY IX.",
        "    MOVE \"AAA\" TO E(1). MOVE \"BBB\" TO E(2). MOVE \"CCC\" TO E(3).\n    MOVE \"DDD\" TO E(4). MOVE \"EEE\" TO E(5).\n    SET IX TO 1.\n    SEARCH E\n        AT END DISPLAY \"NOT FOUND\"\n        WHEN E(IX) = \"AAA\" DISPLAY \"FOUND FIRST\"\n    END-SEARCH.",
    ));
    assert_eq!(out, vec!["FOUND FIRST"]);
}

#[test]
fn search_linear_found_at_last() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC X(3) OCCURS 5 TIMES INDEXED BY IX.",
        "    MOVE \"AAA\" TO E(1). MOVE \"BBB\" TO E(2). MOVE \"CCC\" TO E(3).\n    MOVE \"DDD\" TO E(4). MOVE \"EEE\" TO E(5).\n    SET IX TO 1.\n    SEARCH E\n        AT END DISPLAY \"NOT FOUND\"\n        WHEN E(IX) = \"EEE\" DISPLAY \"FOUND LAST\"\n    END-SEARCH.",
    ));
    assert_eq!(out, vec!["FOUND LAST"]);
}

#[test]
fn search_linear_not_found_at_end() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC X(3) OCCURS 5 TIMES INDEXED BY IX.",
        "    MOVE \"AAA\" TO E(1). MOVE \"BBB\" TO E(2). MOVE \"CCC\" TO E(3).\n    MOVE \"DDD\" TO E(4). MOVE \"EEE\" TO E(5).\n    SET IX TO 1.\n    SEARCH E\n        AT END DISPLAY \"AT END\"\n        WHEN E(IX) = \"ZZZ\" DISPLAY \"FOUND\"\n    END-SEARCH.",
    ));
    assert_eq!(out, vec!["AT END"]);
}

#[test]
fn search_linear_starting_mid_table() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC X(3) OCCURS 5 TIMES INDEXED BY IX.",
        "    MOVE \"AAA\" TO E(1). MOVE \"BBB\" TO E(2). MOVE \"CCC\" TO E(3).\n    MOVE \"DDD\" TO E(4). MOVE \"EEE\" TO E(5).\n    SET IX TO 3.\n    SEARCH E\n        AT END DISPLAY \"NOT FOUND\"\n        WHEN E(IX) = \"CCC\" DISPLAY \"FOUND MID\"\n    END-SEARCH.",
    ));
    assert_eq!(out, vec!["FOUND MID"]);
}

#[test]
fn search_linear_sets_index_on_found() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9(2) OCCURS 5 TIMES INDEXED BY IX.\n01 FOUND-POS PIC 9 VALUE 0.",
        "    MOVE 10 TO E(1). MOVE 20 TO E(2). MOVE 30 TO E(3).\n    MOVE 40 TO E(4). MOVE 50 TO E(5).\n    SET IX TO 1.\n    SEARCH E\n        AT END DISPLAY \"NOT FOUND\"\n        WHEN E(IX) = 30\n            SET FOUND-POS TO IX\n            DISPLAY FOUND-POS\n    END-SEARCH.",
    ));
    assert_eq!(out, vec!["3"]);
}

#[test]
fn search_linear_numeric_table() {
    let out = run_prints(&p(
        "01 T.\n   05 N PIC 9(3) OCCURS 5 TIMES INDEXED BY IX.",
        "    MOVE 100 TO N(1). MOVE 200 TO N(2). MOVE 300 TO N(3).\n    MOVE 400 TO N(4). MOVE 500 TO N(5).\n    SET IX TO 1.\n    SEARCH N\n        AT END DISPLAY \"NOT FOUND\"\n        WHEN N(IX) = 200 DISPLAY \"FOUND 200\"\n    END-SEARCH.",
    ));
    assert_eq!(out, vec!["FOUND 200"]);
}

#[test]
fn search_linear_condition_greater_than() {
    let out = run_prints(&p(
        "01 T.\n   05 N PIC 9(3) OCCURS 5 TIMES INDEXED BY IX.",
        "    MOVE 10 TO N(1). MOVE 20 TO N(2). MOVE 30 TO N(3).\n    MOVE 40 TO N(4). MOVE 50 TO N(5).\n    SET IX TO 1.\n    SEARCH N\n        AT END DISPLAY \"NOT FOUND\"\n        WHEN N(IX) > 25 DISPLAY \"FIRST OVER 25\"\n    END-SEARCH.",
    ));
    assert_eq!(out, vec!["FIRST OVER 25"]);
}

#[test]
fn search_all_multiple_when_condition_compiles() {
    compile_ok(&p(
        "01 SORTED.\n   05 ITEM OCCURS 20 TIMES ASCENDING KEY ITEM-KEY INDEXED BY SI.\n      10 ITEM-KEY PIC 9(4).\n      10 ITEM-VAL PIC X(10).",
        "    SEARCH ALL ITEM\n        AT END\n            DISPLAY \"END\"\n        WHEN ITEM-KEY(SI) = 42\n            DISPLAY ITEM-VAL(SI)\n        WHEN ITEM-KEY(SI) = 99\n            DISPLAY \"NINETY-NINE\"\n    END-SEARCH.",
    ));
}

#[test]
fn search_linear_with_and_condition() {
    compile_ok(&p(
        "01 T.\n   05 E OCCURS 10 TIMES INDEXED BY IX.\n      10 CODE PIC X(2).\n      10 STATUS PIC X VALUE \"A\".",
        "    SET IX TO 1.\n    SEARCH E\n        AT END DISPLAY \"NOT FOUND\"\n        WHEN CODE(IX) = \"AB\" AND STATUS(IX) = \"A\"\n            DISPLAY \"FOUND\"\n    END-SEARCH.",
    ));
}

#[test]
fn search_linear_with_or_condition() {
    compile_ok(&p(
        "01 T.\n   05 E PIC X(3) OCCURS 5 TIMES INDEXED BY IX.",
        "    MOVE \"AAA\" TO E(1). MOVE \"BBB\" TO E(2).\n    SET IX TO 1.\n    SEARCH E\n        AT END DISPLAY \"NOT FOUND\"\n        WHEN E(IX) = \"AAA\" OR E(IX) = \"BBB\"\n            DISPLAY \"FOUND A OR B\"\n    END-SEARCH.",
    ));
}

#[test]
fn search_all_at_end_no_action_body_compiles() {
    compile_ok(&p(
        "01 T.\n   05 V PIC 9(4) OCCURS 10 TIMES ASCENDING KEY V INDEXED BY VI.",
        "    SEARCH ALL V\n        AT END CONTINUE\n        WHEN V(VI) = 7\n            DISPLAY \"FOUND\"\n    END-SEARCH.",
    ));
}

#[test]
fn search_linear_nested_in_loop() {
    compile_ok(&p(
        "01 T.\n   05 E PIC 9(3) OCCURS 5 TIMES INDEXED BY IX.\n01 I PIC 9 VALUE 0.",
        "    MOVE 10 TO E(1). MOVE 20 TO E(2). MOVE 30 TO E(3). MOVE 40 TO E(4). MOVE 50 TO E(5).\n    PERFORM 3 TIMES\n        SET IX TO 1\n        SEARCH E\n            AT END CONTINUE\n            WHEN E(IX) = 30\n                DISPLAY \"30 FOUND\"\n        END-SEARCH\n    END-PERFORM.",
    ));
}

#[test]
fn search_all_two_key_fields_compiles() {
    compile_ok(&p(
        "01 GRADES.\n   05 GRADE OCCURS 30 TIMES ASCENDING KEY STUDENT-ID ASCENDING KEY SUBJECT-ID INDEXED BY GI.\n      10 STUDENT-ID PIC 9(5).\n      10 SUBJECT-ID PIC 9(3).\n      10 SCORE PIC 9(3).",
        "    SEARCH ALL GRADE\n        AT END DISPLAY \"NOT FOUND\"\n        WHEN STUDENT-ID(GI) = 10001 AND SUBJECT-ID(GI) = 101\n            DISPLAY SCORE(GI)\n    END-SEARCH.",
    ));
}

#[test]
fn search_all_performs_action_on_found_field() {
    compile_ok(&p(
        "01 LOOKUP.\n   05 L OCCURS 50 TIMES ASCENDING KEY L-KEY INDEXED BY LI.\n      10 L-KEY PIC 9(5).\n      10 L-VALUE PIC X(20).\n01 FOUND-VAL PIC X(20) VALUE SPACES.",
        "    SEARCH ALL L\n        AT END\n            MOVE \"MISSING\" TO FOUND-VAL\n        WHEN L-KEY(LI) = 12345\n            MOVE L-VALUE(LI) TO FOUND-VAL\n    END-SEARCH.\n    DISPLAY FOUND-VAL.",
    ));
}

#[test]
fn search_linear_multiple_search_calls() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC X(2) OCCURS 5 TIMES INDEXED BY IX.",
        "    MOVE \"A1\" TO E(1). MOVE \"B2\" TO E(2). MOVE \"C3\" TO E(3).\n    MOVE \"D4\" TO E(4). MOVE \"E5\" TO E(5).\n    SET IX TO 1.\n    SEARCH E\n        AT END DISPLAY \"MISS1\"\n        WHEN E(IX) = \"C3\" DISPLAY \"HIT C3\"\n    END-SEARCH.\n    SET IX TO 1.\n    SEARCH E\n        AT END DISPLAY \"MISS2\"\n        WHEN E(IX) = \"E5\" DISPLAY \"HIT E5\"\n    END-SEARCH.",
    ));
    assert_eq!(out, vec!["HIT C3", "HIT E5"]);
}

#[test]
fn search_all_returns_correct_descriptor() {
    compile_ok(&p(
        "01 CATALOG.\n   05 ITEM OCCURS 100 TIMES ASCENDING KEY ITEM-ID INDEXED BY CAT-IDX.\n      10 ITEM-ID PIC 9(6).\n      10 ITEM-NAME PIC X(30).\n      10 ITEM-COST PIC 9(7)V99.\n01 RESULT-NAME PIC X(30) VALUE SPACES.\n01 RESULT-COST PIC 9(7)V99 VALUE 0.",
        "    SEARCH ALL ITEM\n        AT END\n            DISPLAY \"ITEM NOT FOUND\"\n        WHEN ITEM-ID(CAT-IDX) = 200050\n            MOVE ITEM-NAME(CAT-IDX) TO RESULT-NAME\n            MOVE ITEM-COST(CAT-IDX) TO RESULT-COST\n    END-SEARCH.",
    ));
}

#[test]
fn search_linear_action_is_perform() {
    compile_ok(&p(
        "01 T.\n   05 E PIC 9(3) OCCURS 5 TIMES INDEXED BY IX.",
        r#"    MOVE 10 TO E(1). MOVE 20 TO E(2). MOVE 30 TO E(3).
    SET IX TO 1.
    SEARCH E
        AT END DISPLAY "NOT FOUND"
        WHEN E(IX) = 20
            PERFORM FOUND-ACTION
    END-SEARCH.
    STOP RUN.
FOUND-ACTION.
    DISPLAY "ACTION TAKEN"."#,
    ));
}
