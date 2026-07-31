use super::helpers::run_prints;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_set_index_to_integer() {
    let output = run_prints(&p(
        r#"
01 WS-TABLE.
   05 WS-ITEM PIC 9(3) OCCURS 5 TIMES INDEXED BY WS-IDX.
"#,
        r#"
    MOVE 999 TO WS-ITEM(3).
    SET WS-IDX TO 3.
    DISPLAY WS-ITEM(WS-IDX).
"#,
    ));
    assert_eq!(output, vec!["999"]);
}

#[test]
fn test_set_index_to_index() {
    let output = run_prints(&p(
        r#"
01 WS-TABLE.
   05 WS-ITEM PIC 9(3) OCCURS 5 TIMES INDEXED BY WS-IDX1 WS-IDX2.
"#,
        r#"
    MOVE 777 TO WS-ITEM(4).
    SET WS-IDX1 TO 4.
    SET WS-IDX2 TO WS-IDX1.
    DISPLAY WS-ITEM(WS-IDX2).
"#,
    ));
    assert_eq!(output, vec!["777"]);
}

#[test]
fn test_set_index_up_down() {
    let output = run_prints(&p(
        r#"
01 WS-TABLE.
   05 WS-ITEM PIC 9(3) OCCURS 5 TIMES INDEXED BY WS-IDX.
"#,
        r#"
    MOVE 100 TO WS-ITEM(1).
    MOVE 200 TO WS-ITEM(2).
    MOVE 300 TO WS-ITEM(3).
    SET WS-IDX TO 1.
    DISPLAY WS-ITEM(WS-IDX).
    SET WS-IDX UP BY 1.
    DISPLAY WS-ITEM(WS-IDX).
    SET WS-IDX UP BY 1.
    DISPLAY WS-ITEM(WS-IDX).
    SET WS-IDX DOWN BY 2.
    DISPLAY WS-ITEM(WS-IDX).
"#,
    ));
    assert_eq!(output, vec!["100", "200", "300", "100"]);
}

#[test]
fn test_set_pointer() {
    let output = run_prints(&p(
        r#"
01 WS-ITEM PIC X(10).
01 WS-PTR POINTER.
01 WS-FLAG PIC X VALUE 'N'.
"#,
        r#"
    SET WS-PTR TO ADDRESS OF WS-ITEM.
    MOVE 'Y' TO WS-FLAG.
    SET WS-PTR TO NULL.
    IF WS-PTR = NULL
        DISPLAY WS-FLAG
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["Y"]);
}

#[test]
fn test_set_index_with_arithmetic_step() {
    let output = run_prints(&p(
        r#"
01 WS-TABLE.
   05 WS-ITEM PIC 9(2) OCCURS 6 TIMES INDEXED BY WS-IDX.
01 WS-STEP PIC 99 VALUE 2.
"#,
        r#"
    MOVE 10 TO WS-ITEM(1).
    MOVE 20 TO WS-ITEM(3).
    MOVE 30 TO WS-ITEM(5).
    MOVE 1 TO WS-IDX.
    DISPLAY WS-ITEM(WS-IDX).
    SET WS-IDX UP BY WS-STEP.
    DISPLAY WS-ITEM(WS-IDX).
    SET WS-IDX UP BY WS-STEP.
    DISPLAY WS-ITEM(WS-IDX).
"#,
    ));
    assert_eq!(output, vec!["10", "20", "30"]);
}
