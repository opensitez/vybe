use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_odo_structural() {
    compile_ok(&p(
        r#"
01 WS-COUNT PIC 99 VALUE 5.
01 WS-TABLE.
   05 WS-ITEM PIC X(3) OCCURS 1 TO 10 TIMES DEPENDING ON WS-COUNT.
"#,
        r#"
    DISPLAY WS-ITEM(1).
"#,
    ));
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
