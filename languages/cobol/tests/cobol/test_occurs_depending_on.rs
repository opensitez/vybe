use super::helpers::run_prints;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_odo_structural() {
    let output = run_prints(&p(
        r#"
01 WS-COUNT PIC 99 VALUE 2.
01 WS-TABLE.
   05 WS-ITEM PIC X(3) OCCURS 1 TO 10 TIMES DEPENDING ON WS-COUNT.
"#,
        r#"
    MOVE "ONE" TO WS-ITEM(1).
    MOVE "TWO" TO WS-ITEM(2).
    DISPLAY WS-ITEM(1).
    DISPLAY WS-ITEM(2).
"#,
    ));
    assert_eq!(output, vec!["ONE", "TWO"]);
}

#[test]
fn test_odo_varying_limit() {
    let output = run_prints(&p(
        r#"
01 WS-COUNT PIC 99 VALUE 3.
01 WS-TABLE.
   05 WS-ITEM PIC 9(3) OCCURS 1 TO 10 TIMES DEPENDING ON WS-COUNT.
01 WS-I PIC 99 VALUE 1.
"#,
        r#"
    MOVE 100 TO WS-ITEM(1).
    MOVE 200 TO WS-ITEM(2).
    MOVE 300 TO WS-ITEM(3).
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > WS-COUNT
        DISPLAY WS-ITEM(WS-I)
    END-PERFORM.
"#,
    ));
    assert_eq!(output, vec!["100", "200", "300"]);
}

#[test]
fn test_odo_change_count() {
    let output = run_prints(&p(
        r#"
01 WS-COUNT PIC 99 VALUE 2.
01 WS-TABLE.
   05 WS-ITEM PIC 9(3) OCCURS 1 TO 10 TIMES DEPENDING ON WS-COUNT.
"#,
        r#"
    MOVE 111 TO WS-ITEM(1).
    MOVE 222 TO WS-ITEM(2).
    DISPLAY WS-ITEM(1).
    DISPLAY WS-ITEM(2).
    MOVE 3 TO WS-COUNT.
    MOVE 333 TO WS-ITEM(3).
    DISPLAY WS-ITEM(3).
"#,
    ));
    assert_eq!(output, vec!["111", "222", "333"]);
}

#[test]
fn test_odo_boundaries_minimum_count() {
    let output = run_prints(&p(
        r#"
01 WS-COUNT PIC 99 VALUE 1.
01 WS-TABLE.
   05 WS-ITEM PIC 9(3) OCCURS 1 TO 10 TIMES DEPENDING ON WS-COUNT.
01 WS-I PIC 99 VALUE 1.
"#,
        r#"
    MOVE 101 TO WS-ITEM(1).
    MOVE 102 TO WS-ITEM(2).
    MOVE 103 TO WS-ITEM(3).
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > WS-COUNT
        DISPLAY WS-ITEM(WS-I)
    END-PERFORM.
"#,
    ));
    assert_eq!(output, vec!["101"]);
}

#[test]
fn test_odo_boundaries_maximum_count() {
    let output = run_prints(&p(
        r#"
01 WS-COUNT PIC 99 VALUE 5.
01 WS-TABLE.
   05 WS-ITEM PIC 9(3) OCCURS 1 TO 5 TIMES DEPENDING ON WS-COUNT.
01 WS-I PIC 99 VALUE 1.
"#,
        r#"
    MOVE 1 TO WS-ITEM(1).
    MOVE 2 TO WS-ITEM(2).
    MOVE 3 TO WS-ITEM(3).
    MOVE 4 TO WS-ITEM(4).
    MOVE 5 TO WS-ITEM(5).
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > WS-COUNT
        DISPLAY WS-ITEM(WS-I)
    END-PERFORM.
"#,
    ));
    assert_eq!(output, vec!["1", "2", "3", "4", "5"]);
}

#[test]
fn test_odo_shrink_then_expand() {
    let output = run_prints(&p(
        r#"
01 WS-COUNT PIC 99 VALUE 4.
01 WS-TABLE.
   05 WS-ITEM PIC 9(3) OCCURS 1 TO 10 TIMES DEPENDING ON WS-COUNT.
01 WS-I PIC 99 VALUE 1.
"#,
        r#"
    MOVE 10 TO WS-ITEM(1).
    MOVE 20 TO WS-ITEM(2).
    MOVE 30 TO WS-ITEM(3).
    MOVE 40 TO WS-ITEM(4).
    MOVE 4 TO WS-COUNT.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > WS-COUNT
        DISPLAY WS-ITEM(WS-I)
    END-PERFORM.
    MOVE 2 TO WS-COUNT.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > WS-COUNT
        DISPLAY WS-ITEM(WS-I)
    END-PERFORM.
    MOVE 3 TO WS-COUNT.
    MOVE 55 TO WS-ITEM(3).
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > WS-COUNT
        DISPLAY WS-ITEM(WS-I)
    END-PERFORM.
"#,
    ));
    assert_eq!(
        output,
        vec!["10", "20", "30", "40", "10", "20", "10", "20", "55"]
    );
}

#[test]
fn test_odo_with_reference_modification() {
    let output = run_prints(&p(
        r#"
01 WS-COUNT PIC 99 VALUE 2.
01 WS-TABLE.
   05 WS-ITEM PIC X(4) OCCURS 1 TO 5 TIMES DEPENDING ON WS-COUNT.
01 WS-SUB PIC 99 VALUE 1.
"#,
        r#"
    MOVE "ABCD" TO WS-ITEM(1).
    MOVE "EFGH" TO WS-ITEM(2).
    DISPLAY WS-ITEM(WS-SUB)(2:2).
    DISPLAY WS-ITEM(WS-SUB + 1)(3:2).
"#,
    ));
    assert_eq!(output, vec!["BC", "GH"]);
}

#[test]
fn test_odo_indexed_lookup() {
    let output = run_prints(&p(
        r#"
01 WS-COUNT PIC 99 VALUE 3.
01 WS-TABLE.
   05 WS-ENTRY PIC 9(2) OCCURS 1 TO 6 TIMES DEPENDING ON WS-COUNT INDEXED BY WS-IDX.
"#,
        r#"
    MOVE 12 TO WS-ENTRY(1).
    MOVE 24 TO WS-ENTRY(2).
    MOVE 36 TO WS-ENTRY(3).
    MOVE 48 TO WS-ENTRY(4).
    SET WS-IDX TO 2.
    DISPLAY WS-ENTRY(WS-IDX).
"#,
    ));
    assert_eq!(output, vec!["24"]);
}

#[test]
fn test_odo_search_like_pattern_not_allowed() {
    let output = run_prints(&p(
        r#"
01 WS-COUNT PIC 99 VALUE 3.
01 WS-TABLE.
   05 WS-ENTRY OCCURS 1 TO 6 TIMES DEPENDING ON WS-COUNT.
      10 WS-VAL PIC 9(2).
01 WS-I PIC 99 VALUE 1.
"#,
        r#"
    MOVE 11 TO WS-VAL(1).
    MOVE 22 TO WS-VAL(2).
    MOVE 33 TO WS-VAL(3).
    MOVE 11 TO WS-VAL(WS-I).
    DISPLAY WS-VAL(WS-I).
"#,
    ));
    assert_eq!(output, vec!["11"]);
}

#[test]
fn test_odo_indexed_ascending_loop() {
    let output = run_prints(&p(
        r#"
01 WS-COUNT PIC 99 VALUE 3.
01 WS-TABLE.
   05 WS-ENTRY PIC 9(2) OCCURS 1 TO 6 TIMES DEPENDING ON WS-COUNT INDEXED BY WS-IDX.
01 WS-IDX PIC 99 VALUE 1.
"#,
        r#"
    MOVE 11 TO WS-ENTRY(1).
    MOVE 22 TO WS-ENTRY(2).
    MOVE 33 TO WS-ENTRY(3).
    MOVE 44 TO WS-ENTRY(4).
    SET WS-IDX TO 1.
    PERFORM UNTIL WS-IDX > WS-COUNT
        DISPLAY WS-ENTRY(WS-IDX)
        SET WS-IDX UP BY 1
    END-PERFORM.
"#,
    ));
    assert_eq!(output, vec!["11", "22", "33"]);
}

#[test]
fn test_odo_indexed_descending_loop() {
    let output = run_prints(&p(
        r#"
01 WS-COUNT PIC 99 VALUE 4.
01 WS-TABLE.
   05 WS-ENTRY PIC 9(2) OCCURS 1 TO 6 TIMES DEPENDING ON WS-COUNT INDEXED BY WS-IDX.
01 WS-IDX PIC 99 VALUE 1.
"#,
        r#"
    MOVE 11 TO WS-ENTRY(1).
    MOVE 22 TO WS-ENTRY(2).
    MOVE 33 TO WS-ENTRY(3).
    MOVE 44 TO WS-ENTRY(4).
    SET WS-IDX TO WS-COUNT.
    PERFORM UNTIL WS-IDX < 1
        DISPLAY WS-ENTRY(WS-IDX)
        SET WS-IDX DOWN BY 1
    END-PERFORM.
"#,
    ));
    assert_eq!(output, vec!["44", "33", "22", "11"]);
}

#[test]
fn test_odo_recount_with_addition() {
    let output = run_prints(&p(
        r#"
01 WS-COUNT PIC 99 VALUE 2.
01 WS-TABLE.
   05 WS-ENTRY PIC 9(2) OCCURS 1 TO 6 TIMES DEPENDING ON WS-COUNT.
01 WS-I PIC 99 VALUE 1.
"#,
        r#"
    MOVE 10 TO WS-ENTRY(1).
    MOVE 20 TO WS-ENTRY(2).
    MOVE 30 TO WS-ENTRY(3).
    MOVE 40 TO WS-ENTRY(4).
    ADD 2 TO WS-COUNT.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > WS-COUNT
        DISPLAY WS-ENTRY(WS-I)
    END-PERFORM.
"#,
    ));
    assert_eq!(output, vec!["10", "20", "30", "40"]);
}

#[test]
fn test_odo_recount_with_subtraction() {
    let output = run_prints(&p(
        r#"
01 WS-COUNT PIC 99 VALUE 4.
01 WS-TABLE.
   05 WS-ENTRY PIC 9(2) OCCURS 1 TO 6 TIMES DEPENDING ON WS-COUNT.
01 WS-I PIC 99 VALUE 1.
"#,
        r#"
    MOVE 10 TO WS-ENTRY(1).
    MOVE 20 TO WS-ENTRY(2).
    MOVE 30 TO WS-ENTRY(3).
    MOVE 40 TO WS-ENTRY(4).
    SUBTRACT 2 FROM WS-COUNT.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > WS-COUNT
        DISPLAY WS-ENTRY(WS-I)
    END-PERFORM.
"#,
    ));
    assert_eq!(output, vec!["10", "20"]);
}

#[test]
fn test_odo_count_resize_cycles() {
    let output = run_prints(&p(
        r#"
01 WS-COUNT PIC 99 VALUE 2.
01 WS-TABLE.
   05 WS-ENTRY PIC 9(2) OCCURS 1 TO 6 TIMES DEPENDING ON WS-COUNT.
01 WS-I PIC 99 VALUE 1.
"#,
        r#"
    MOVE 11 TO WS-ENTRY(1).
    MOVE 22 TO WS-ENTRY(2).
    MOVE 33 TO WS-ENTRY(3).
    MOVE 44 TO WS-ENTRY(4).
    MOVE 55 TO WS-ENTRY(5).
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > WS-COUNT
        DISPLAY WS-ENTRY(WS-I)
    END-PERFORM.
    MOVE 1 TO WS-COUNT.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > WS-COUNT
        DISPLAY WS-ENTRY(WS-I)
    END-PERFORM.
    MOVE 4 TO WS-COUNT.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > WS-COUNT
        DISPLAY WS-ENTRY(WS-I)
    END-PERFORM.
"#,
    ));
    assert_eq!(output, vec!["11", "22", "11", "11", "22", "33", "44"]);
}

#[test]
fn test_odo_two_tables_share_dep_on_count() {
    let output = run_prints(&p(
        r#"
01 WS-COUNT PIC 99 VALUE 2.
01 WS-TABLE-STR.
   05 WS-LABEL PIC X(2) OCCURS 1 TO 5 TIMES DEPENDING ON WS-COUNT.
01 WS-TABLE-NUM.
   05 WS-NUMBER PIC 9(2) OCCURS 1 TO 5 TIMES DEPENDING ON WS-COUNT.
01 WS-I PIC 99 VALUE 1.
"#,
        r#"
    MOVE "AA" TO WS-LABEL(1).
    MOVE "BB" TO WS-LABEL(2).
    MOVE "CC" TO WS-LABEL(3).
    MOVE 10 TO WS-NUMBER(1).
    MOVE 20 TO WS-NUMBER(2).
    MOVE 30 TO WS-NUMBER(3).
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > WS-COUNT
        DISPLAY WS-LABEL(WS-I)
        DISPLAY WS-NUMBER(WS-I)
    END-PERFORM.
    MOVE 3 TO WS-COUNT.
    MOVE "CC" TO WS-LABEL(3).
    MOVE 30 TO WS-NUMBER(3).
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > WS-COUNT
        DISPLAY WS-LABEL(WS-I)
        DISPLAY WS-NUMBER(WS-I)
    END-PERFORM.
"#,
    ));
    assert_eq!(
        output,
        vec!["AA", "10", "BB", "20", "AA", "10", "BB", "20", "CC", "30"],
    );
}

#[test]
fn test_odo_indexed_step_scan() {
    let output = run_prints(&p(
        r#"
01 WS-COUNT PIC 99 VALUE 4.
01 WS-TABLE.
   05 WS-ENTRY PIC 9(2) OCCURS 1 TO 4 TIMES DEPENDING ON WS-COUNT.
01 WS-I PIC 99 VALUE 1.
"#,
        r#"
    MOVE 10 TO WS-ENTRY(1).
    MOVE 20 TO WS-ENTRY(2).
    MOVE 30 TO WS-ENTRY(3).
    MOVE 40 TO WS-ENTRY(4).
    MOVE 1 TO WS-I.
    PERFORM UNTIL WS-I > WS-COUNT
        DISPLAY WS-ENTRY(WS-I)
        ADD 2 TO WS-I
    END-PERFORM.
"#,
    ));
    assert_eq!(output, vec!["10", "30"]);
}

#[test]
fn test_odo_reference_mod_with_active_count() {
    let output = run_prints(&p(
        r#"
01 WS-COUNT PIC 99 VALUE 2.
01 WS-TABLE.
   05 WS-ENTRY PIC X(4) OCCURS 1 TO 5 TIMES DEPENDING ON WS-COUNT.
"#,
        r#"
    MOVE "ABCD" TO WS-ENTRY(1).
    MOVE "EFGH" TO WS-ENTRY(2).
    MOVE "IJKL" TO WS-ENTRY(3).
    DISPLAY WS-ENTRY(WS-COUNT)(2:2).
    MOVE 1 TO WS-COUNT.
    DISPLAY WS-ENTRY(WS-COUNT)(2:2).
"#,
    ));
    assert_eq!(output, vec!["FG", "BC"]);
}

#[test]
fn test_odo_access_outside_active_count() {
    let output = run_prints(&p(
        r#"
01 WS-COUNT PIC 99 VALUE 2.
01 WS-TABLE.
   05 WS-ENTRY PIC 9(2) OCCURS 1 TO 6 TIMES DEPENDING ON WS-COUNT.
01 WS-I PIC 99 VALUE 1.
"#,
        r#"
    MOVE 11 TO WS-ENTRY(1).
    MOVE 22 TO WS-ENTRY(2).
    MOVE 33 TO WS-ENTRY(3).
    MOVE 44 TO WS-ENTRY(4).
    DISPLAY WS-ENTRY(4).
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > WS-COUNT
        DISPLAY WS-ENTRY(WS-I)
    END-PERFORM.
"#,
    ));
    assert_eq!(output, vec!["44", "11", "22"]);
}

#[test]
fn test_odo_references_still_hold_values_after_count_change() {
    let output = run_prints(&p(
        r#"
01 WS-COUNT PIC 99 VALUE 4.
01 WS-TABLE.
   05 WS-ENTRY PIC 9(3) OCCURS 1 TO 6 TIMES DEPENDING ON WS-COUNT.
"#,
        r#"
    MOVE 101 TO WS-ENTRY(1).
    MOVE 102 TO WS-ENTRY(2).
    MOVE 103 TO WS-ENTRY(3).
    MOVE 104 TO WS-ENTRY(4).
    MOVE 2 TO WS-COUNT.
    DISPLAY WS-ENTRY(4).
    MOVE 205 TO WS-ENTRY(4).
    DISPLAY WS-ENTRY(4).
"#,
    ));
    assert_eq!(output, vec!["104", "205"]);
}
