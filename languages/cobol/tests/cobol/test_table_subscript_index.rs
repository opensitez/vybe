use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn table_subscript_integer_literal() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9(2) OCCURS 5 TIMES.",
        "    MOVE 77 TO E(3).\n    DISPLAY E(3).",
    ));
    assert_eq!(out, vec!["77"]);
}

#[test]
fn table_subscript_variable() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9(2) OCCURS 5 TIMES.\n01 I PIC 9 VALUE 2.",
        "    MOVE 42 TO E(I).\n    DISPLAY E(I).",
    ));
    assert_eq!(out, vec!["42"]);
}

#[test]
fn table_fill_all_elements() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9 OCCURS 4 TIMES.\n01 I PIC 9 VALUE 0.",
        "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 4\n        MOVE I TO E(I)\n    END-PERFORM.\n    DISPLAY E(1).\n    DISPLAY E(4).",
    ));
    assert_eq!(out, vec!["1", "4"]);
}

#[test]
fn table_element_add() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9(3) OCCURS 3 TIMES.",
        "    MOVE 100 TO E(1).\n    ADD 50 TO E(1).\n    DISPLAY E(1).",
    ));
    assert_eq!(out, vec!["150"]);
}

#[test]
fn table_element_subtract() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9(3) OCCURS 3 TIMES.",
        "    MOVE 200 TO E(2).\n    SUBTRACT 75 FROM E(2).\n    DISPLAY E(2).",
    ));
    assert_eq!(out, vec!["125"]);
}

#[test]
fn table_element_multiply() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9(4) OCCURS 3 TIMES.",
        "    MOVE 12 TO E(3).\n    MULTIPLY 4 BY E(3).\n    DISPLAY E(3).",
    ));
    assert_eq!(out, vec!["48"]);
}

#[test]
fn table_sum_all_elements() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9(3) OCCURS 5 TIMES.\n01 S PIC 9(5) VALUE 0.\n01 I PIC 9 VALUE 0.",
        "    MOVE 10 TO E(1). MOVE 20 TO E(2). MOVE 30 TO E(3).\n    MOVE 40 TO E(4). MOVE 50 TO E(5).\n    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 5\n        ADD E(I) TO S\n    END-PERFORM.\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["00150"]);
}

#[test]
fn table_element_comparison() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9(3) OCCURS 3 TIMES.",
        "    MOVE 100 TO E(1).\n    MOVE 200 TO E(2).\n    IF E(2) > E(1)\n        DISPLAY \"SECOND BIGGER\"\n    ELSE\n        DISPLAY \"FIRST BIGGER OR EQUAL\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["SECOND BIGGER"]);
}

#[test]
fn table_two_dim_access() {
    let out = run_prints(&p(
        "01 M.\n   05 ROW OCCURS 3 TIMES.\n      10 COL PIC 9 OCCURS 3 TIMES.",
        "    MOVE 5 TO COL(2 2).\n    DISPLAY COL(2 2).",
    ));
    assert_eq!(out, vec!["5"]);
}

#[test]
fn table_two_dim_sum_main_diagonal() {
    let out = run_prints(&p(
        "01 M.\n   05 ROW OCCURS 3 TIMES.\n      10 COL PIC 9 OCCURS 3 TIMES.\n01 S PIC 9(3) VALUE 0.\n01 I PIC 9 VALUE 0.",
        "    MOVE 1 TO COL(1 1). MOVE 2 TO COL(2 2). MOVE 3 TO COL(3 3).\n    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 3\n        ADD COL(I I) TO S\n    END-PERFORM.\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["6"]);
}

#[test]
fn table_element_move_to_ws() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC X(5) OCCURS 3 TIMES.\n01 COPY PIC X(5) VALUE SPACES.",
        "    MOVE \"HELLO\" TO E(2).\n    MOVE E(2) TO COPY.\n    DISPLAY COPY.",
    ));
    assert_eq!(out, vec!["HELLO"]);
}

#[test]
fn table_search_found_second_element() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC X(3) OCCURS 5 TIMES INDEXED BY IX.",
        "    MOVE \"AAA\" TO E(1). MOVE \"BBB\" TO E(2). MOVE \"CCC\" TO E(3).\n    MOVE \"DDD\" TO E(4). MOVE \"EEE\" TO E(5).\n    SET IX TO 1.\n    SEARCH E\n        AT END DISPLAY \"MISS\"\n        WHEN E(IX) = \"BBB\" DISPLAY \"HIT\"\n    END-SEARCH.",
    ));
    assert_eq!(out, vec!["HIT"]);
}

#[test]
fn table_subscript_arithmetic_expression() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9 OCCURS 5 TIMES.\n01 I PIC 9 VALUE 2.",
        "    MOVE 7 TO E(I + 1).\n    DISPLAY E(3).",
    ));
    assert_eq!(out, vec!["7"]);
}

#[test]
fn table_max_element_find() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9(3) OCCURS 5 TIMES.\n01 MAX PIC 9(3) VALUE 0.\n01 I PIC 9 VALUE 0.",
        "    MOVE 30 TO E(1). MOVE 70 TO E(2). MOVE 50 TO E(3).\n    MOVE 90 TO E(4). MOVE 10 TO E(5).\n    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 5\n        IF E(I) > MAX\n            MOVE E(I) TO MAX\n        END-IF\n    END-PERFORM.\n    DISPLAY MAX.",
    ));
    assert_eq!(out, vec!["090"]);
}

#[test]
fn table_element_in_evaluate() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC X OCCURS 3 TIMES.",
        "    MOVE \"B\" TO E(2).\n    EVALUATE E(2)\n        WHEN \"A\" DISPLAY \"ALPHA\"\n        WHEN \"B\" DISPLAY \"BETA\"\n        WHEN OTHER DISPLAY \"OTHER\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["BETA"]);
}

#[test]
fn table_string_table_initialized_and_displayed() {
    let out = run_prints(&p(
        "01 WORDS.\n   05 WORD PIC X(5) OCCURS 3 TIMES.",
        "    MOVE \"ONE  \" TO WORD(1).\n    MOVE \"TWO  \" TO WORD(2).\n    MOVE \"THREE\" TO WORD(3).\n    DISPLAY WORD(1).\n    DISPLAY WORD(2).\n    DISPLAY WORD(3).",
    ));
    assert_eq!(out, vec!["ONE  ", "TWO  ", "THREE"]);
}

#[test]
fn table_reference_modification_in_element() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC X(6) OCCURS 3 TIMES.",
        "    MOVE \"ABCDEF\" TO E(1).\n    DISPLAY E(1)(2:3).",
    ));
    assert_eq!(out, vec!["BCD"]);
}

#[test]
fn table_element_numeric_picture_displays_zeroes() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9(4) OCCURS 3 TIMES.",
        "    DISPLAY E(2).",
    ));
    assert_eq!(out, vec!["0000"]);
}

#[test]
fn table_element_count_occurrences() {
    let out = run_prints(&p(
        "01 T.\n   05 GRADE PIC X OCCURS 10 TIMES.\n01 CNT PIC 9(2) VALUE 0.\n01 I PIC 9(2) VALUE 0.",
        "    MOVE \"A\" TO GRADE(1). MOVE \"B\" TO GRADE(2). MOVE \"A\" TO GRADE(3).\n    MOVE \"C\" TO GRADE(4). MOVE \"A\" TO GRADE(5).\n    MOVE \"B\" TO GRADE(6). MOVE \"D\" TO GRADE(7). MOVE \"A\" TO GRADE(8).\n    MOVE \"F\" TO GRADE(9). MOVE \"A\" TO GRADE(10).\n    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 10\n        IF GRADE(I) = \"A\"\n            ADD 1 TO CNT\n        END-IF\n    END-PERFORM.\n    DISPLAY CNT.",
    ));
    assert_eq!(out, vec!["05"]);
}

#[test]
fn table_copy_one_element_to_another() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9(3) OCCURS 5 TIMES.",
        "    MOVE 123 TO E(1).\n    MOVE E(1) TO E(5).\n    DISPLAY E(5).",
    ));
    assert_eq!(out, vec!["123"]);
}

#[test]
fn table_sequential_search_not_found() {
    let out = run_prints(&p(
        "01 T.\n   05 KEY PIC X(3) OCCURS 5 TIMES INDEXED BY IX.",
        "    MOVE \"AAA\" TO KEY(1). MOVE \"BBB\" TO KEY(2). MOVE \"CCC\" TO KEY(3).\n    MOVE \"DDD\" TO KEY(4). MOVE \"EEE\" TO KEY(5).\n    SET IX TO 1.\n    SEARCH KEY\n        AT END DISPLAY \"NOT FOUND\"\n        WHEN KEY(IX) = \"ZZZ\"\n            DISPLAY \"FOUND\"\n    END-SEARCH.",
    ));
    assert_eq!(out, vec!["NOT FOUND"]);
}

#[test]
fn table_compute_using_subscripted_field() {
    let out = run_prints(&p(
        "01 T.\n   05 N PIC 9(3) OCCURS 3 TIMES.\n01 R PIC 9(5) VALUE 0.",
        "    MOVE 5 TO N(1). MOVE 10 TO N(2). MOVE 15 TO N(3).\n    COMPUTE R = N(1) * N(2) + N(3).\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["65"]);
}

#[test]
fn table_element_in_string_concat() {
    compile_ok(&p(
        "01 NAMES.\n   05 NAME PIC X(5) OCCURS 3 TIMES.\n01 RESULT PIC X(20).",
        "    MOVE \"ALICE\" TO NAME(1).\n    MOVE \"BOB  \" TO NAME(2).\n    STRING NAME(1) DELIMITED BY SPACE \" \" DELIMITED BY SIZE\n           NAME(2) DELIMITED BY SPACE INTO RESULT.",
    ));
}

#[test]
fn table_ascending_initialized_order() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9(3) OCCURS 5 TIMES.\n01 I PIC 9 VALUE 0.",
        "    MOVE 10 TO E(1). MOVE 20 TO E(2). MOVE 30 TO E(3).\n    MOVE 40 TO E(4). MOVE 50 TO E(5).\n    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 5\n        DISPLAY E(I)\n    END-PERFORM.",
    ));
    assert_eq!(out, vec!["010", "020", "030", "040", "050"]);
}

#[test]
fn table_binary_search_found() {
    compile_ok(&p(
        "01 SORTED-T.\n   05 ENTRY OCCURS 10 TIMES ASCENDING KEY ENTRY INDEXED BY S-IX.\n      10 ENTRY PIC 9(4).",
        "    SEARCH ALL ENTRY\n        AT END DISPLAY \"NOT FOUND\"\n        WHEN ENTRY(S-IX) = 5\n            DISPLAY \"FOUND\"\n    END-SEARCH.",
    ));
}

#[test]
fn table_binary_search_at_end() {
    compile_ok(&p(
        "01 SORTED-T.\n   05 ENTRY OCCURS 5 TIMES ASCENDING KEY ENTRY INDEXED BY S-IX.\n      10 ENTRY PIC 9(2).",
        "    SEARCH ALL ENTRY\n        AT END\n            DISPLAY \"NOT IN TABLE\"\n        WHEN ENTRY(S-IX) = 99\n            DISPLAY \"FOUND 99\"\n    END-SEARCH.",
    ));
}
