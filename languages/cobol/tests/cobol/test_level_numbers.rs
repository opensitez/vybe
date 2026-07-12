use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_level_nesting_basics() {
    let output = run_prints(&p(
        r#"
01 WS-GROUP.
   05 WS-SUB1 PIC X(3) VALUE "ABC".
   05 WS-SUB2 PIC X(3) VALUE "DEF".
"#,
        r#"
    DISPLAY WS-GROUP.
    DISPLAY WS-SUB1.
    DISPLAY WS-SUB2.
"#,
    ));
    assert_eq!(output, vec!["ABCDEF", "ABC", "DEF"]);
}

#[test]
fn test_level_three_nesting() {
    let output = run_prints(&p(
        r#"
01 WS-TOP.
   05 WS-MID.
      10 WS-BOT PIC X(3) VALUE "XYZ".
"#,
        r#"
    DISPLAY WS-TOP.
    DISPLAY WS-MID.
    DISPLAY WS-BOT.
"#,
    ));
    assert_eq!(output, vec!["XYZ", "XYZ", "XYZ"]);
}

#[test]
fn test_level_deep_nesting() {
    compile_ok(&p(
        r#"
01 WS-L01.
   05 WS-L05.
      10 WS-L10.
         15 WS-L15.
            20 WS-L20 PIC X(5) VALUE "HELLO".
"#,
        r#"
    DISPLAY WS-L01.
"#,
    ));
}

#[test]
fn test_level_77_item() {
    let output = run_prints(&p(
        r#"
77 WS-INT PIC 9(3) VALUE 100.
77 WS-STR PIC X(5) VALUE "HELLO".
"#,
        r#"
    ADD 50 TO WS-INT.
    DISPLAY WS-INT.
    DISPLAY WS-STR.
"#,
    ));
    assert_eq!(output, vec!["150", "HELLO"]);
}

#[test]
fn test_level_88_single_value() {
    let output = run_prints(&p(
        r#"
01 WS-STATUS PIC 9 VALUE 1.
   88 IS-ACTIVE VALUE 1.
   88 IS-INACTIVE VALUE 0.
"#,
        r#"
    IF IS-ACTIVE
        DISPLAY "ACTIVE"
    ELSE
        DISPLAY "INACTIVE"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["ACTIVE"]);
}

#[test]
fn test_level_88_multiple_values() {
    let output = run_prints(&p(
        r#"
01 WS-CODE PIC X VALUE "B".
   88 IS-VALID-CODE VALUE "A", "B", "C".
"#,
        r#"
    IF IS-VALID-CODE
        DISPLAY "VALID"
    ELSE
        DISPLAY "INVALID"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["VALID"]);
}

#[test]
fn test_level_88_value_thru() {
    let output = run_prints(&p(
        r#"
01 WS-AGE PIC 9(3) VALUE 25.
   88 IS-YOUTH VALUE 15 THRU 30.
"#,
        r#"
    IF IS-YOUTH
        DISPLAY "YOUTH"
    ELSE
        DISPLAY "OTHER"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["YOUTH"]);
}

#[test]
fn test_level_88_set_to_true() {
    let output = run_prints(&p(
        r#"
01 WS-FLAG PIC 9 VALUE 0.
   88 IS-ON VALUE 1.
"#,
        r#"
    SET IS-ON TO TRUE.
    DISPLAY WS-FLAG.
"#,
    ));
    assert_eq!(output, vec!["1"]);
}

#[test]
fn test_level_filler_items() {
    let output = run_prints(&p(
        r#"
01 WS-RECORD.
   05 WS-FIELD1 PIC X(3) VALUE "AAA".
   05 FILLER PIC X(3) VALUE "BBB".
   05 WS-FIELD2 PIC X(3) VALUE "CCC".
"#,
        r#"
    DISPLAY WS-RECORD.
"#,
    ));
    assert_eq!(output, vec!["AAABBBCCC"]);
}

#[test]
fn test_level_siblings_same_depth() {
    let output = run_prints(&p(
        r#"
01 WS-RECORD.
   05 WS-A PIC X(3) VALUE "AAA".
   05 WS-B PIC X(3) VALUE "BBB".
   05 WS-C PIC X(3) VALUE "CCC".
"#,
        r#"
    DISPLAY WS-RECORD.
"#,
    ));
    assert_eq!(output, vec!["AAABBBCCC"]);
}
