use super::helpers::run_prints;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_table_occurs_keys() {
    let output = run_prints(&p(
        r#"
01 WS-TABLE.
   05 WS-ENTRY OCCURS 5 TIMES
      ASCENDING KEY IS WS-KEY
      INDEXED BY WS-IDX.
      10 WS-KEY PIC 9(3).
"#,
        r#"
    MOVE 10 TO WS-KEY(1).
    MOVE 20 TO WS-KEY(2).
    MOVE 30 TO WS-KEY(3).
    SEARCH ALL WS-ENTRY
        WHEN WS-KEY(WS-IDX) = 20 DISPLAY 'FOUND'
    END-SEARCH.
"#,
    ));
    assert_eq!(output, vec!["FOUND"]);
}

#[test]
fn test_table_occurs_keys_not_found() {
    let output = run_prints(&p(
        r#"
01 WS-TABLE.
   05 WS-ENTRY OCCURS 4 TIMES
      ASCENDING KEY IS WS-KEY
      INDEXED BY WS-IDX.
      10 WS-KEY PIC 9(3).
"#,
        r#"
    MOVE 10 TO WS-KEY(1).
    MOVE 20 TO WS-KEY(2).
    MOVE 30 TO WS-KEY(3).
    SEARCH ALL WS-ENTRY
        AT END DISPLAY 'NONE'
        WHEN WS-KEY(WS-IDX) = 99 DISPLAY 'FOUND'
    END-SEARCH.
"#,
    ));
    assert_eq!(output, vec!["NONE"]);
}

#[test]
fn test_table_linear_search_with_at_end() {
    let output = run_prints(&p(
        r#"
01 WS-TABLE.
   05 WS-ENTRY OCCURS 4 TIMES
      INDEXED BY WS-IDX.
      10 WS-KEY PIC X(2).
"#,
        r#"
    MOVE "A" TO WS-KEY(1).
    MOVE "B" TO WS-KEY(2).
    MOVE "C" TO WS-KEY(3).
    SEARCH WS-ENTRY
        AT END DISPLAY 'NONE'
        WHEN WS-KEY(WS-IDX) = "B" DISPLAY 'FOUND'
    END-SEARCH.
"#,
    ));
    assert_eq!(output, vec!["FOUND"]);
}

#[test]
fn test_table_2d_occurs() {
    let output = run_prints(&p(
        r#"
01 WS-TABLE-2D.
   05 WS-ROW OCCURS 3 TIMES.
      10 WS-COL OCCURS 3 TIMES.
         15 WS-CELL PIC 9 VALUE 5.
"#,
        r#"
    DISPLAY WS-CELL(2 3).
"#,
    ));
    assert_eq!(output, vec!["5"]);
}

#[test]
fn test_table_index_access() {
    let output = run_prints(&p(
        r#"
01 WS-TABLE.
   05 WS-ITEM PIC 9(3) OCCURS 5 TIMES INDEXED BY WS-IDX.
"#,
        r#"
    SET WS-IDX TO 1.
    MOVE 123 TO WS-ITEM(WS-IDX).
    SET WS-IDX UP BY 1.
    MOVE 456 TO WS-ITEM(WS-IDX).
    SET WS-IDX DOWN BY 1.
    DISPLAY WS-ITEM(WS-IDX).
    SET WS-IDX UP BY 1.
    DISPLAY WS-ITEM(WS-IDX).
"#,
    ));
    assert_eq!(output, vec!["123", "456"]);
}

#[test]
fn test_table_subscript_variable() {
    let output = run_prints(&p(
        r#"
01 WS-TABLE.
   05 WS-ITEM PIC 9(3) OCCURS 5 TIMES.
01 WS-SUB PIC 9 VALUE 3.
"#,
        r#"
    MOVE 999 TO WS-ITEM(WS-SUB).
    DISPLAY WS-ITEM(WS-SUB).
"#,
    ));
    assert_eq!(output, vec!["999"]);
}

#[test]
fn test_table_subscript_expression() {
    let output = run_prints(&p(
        r#"
01 WS-TABLE.
   05 WS-ITEM PIC 9(3) OCCURS 5 TIMES.
01 WS-SUB PIC 9 VALUE 2.
"#,
        r#"
    MOVE 777 TO WS-ITEM(WS-SUB + 1).
    DISPLAY WS-ITEM(WS-SUB + 1).
"#,
    ));
    assert_eq!(output, vec!["777"]);
}

#[test]
fn test_table_move_zeros() {
    let output = run_prints(&p(
        r#"
01 WS-TABLE.
   05 WS-ITEM PIC 9(3) OCCURS 3 TIMES VALUE 111.
"#,
        r#"
    MOVE ZEROS TO WS-TABLE.
    DISPLAY WS-ITEM(1).
    DISPLAY WS-ITEM(2).
    DISPLAY WS-ITEM(3).
"#,
    ));
    assert_eq!(output, vec!["000", "000", "000"]);
}

#[test]
fn test_table_copy_loop() {
    let output = run_prints(&p(
        r#"
01 WS-TABLE1.
   05 WS-ITEM1 PIC X(3) OCCURS 3 TIMES.
01 WS-TABLE2.
   05 WS-ITEM2 PIC X(3) OCCURS 3 TIMES.
01 WS-I PIC 9 VALUE 1.
"#,
        r#"
    MOVE "AAA" TO WS-ITEM1(1).
    MOVE "BBB" TO WS-ITEM1(2).
    MOVE "CCC" TO WS-ITEM1(3).
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 3
        MOVE WS-ITEM1(WS-I) TO WS-ITEM2(WS-I)
    END-PERFORM.
    DISPLAY WS-ITEM2(1).
    DISPLAY WS-ITEM2(2).
    DISPLAY WS-ITEM2(3).
"#,
    ));
    assert_eq!(output, vec!["AAA", "BBB", "CCC"]);
}

#[test]
fn test_table_refmod_element() {
    let output = run_prints(&p(
        r#"
01 WS-TABLE.
   05 WS-ITEM PIC X(5) OCCURS 3 TIMES.
"#,
        r#"
    MOVE "HELLO" TO WS-ITEM(2).
    DISPLAY WS-ITEM(2)(1:3).
"#,
    ));
    assert_eq!(output, vec!["HEL"]);
}
