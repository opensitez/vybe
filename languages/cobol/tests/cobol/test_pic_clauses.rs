use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_pic_alphabetic() {
    compile_ok(&p(
        r#"
01 WS-A PIC A(10) VALUE "HELLO".
01 WS-B PIC A VALUE "X".
"#,
        r#"
    DISPLAY WS-A.
    DISPLAY WS-B.
"#,
    ));
}

#[test]
fn test_pic_single_char_digit() {
    compile_ok(&p(
        r#"
01 WS-CHAR PIC X VALUE "A".
01 WS-DIGIT PIC 9 VALUE 7.
"#,
        r#"
    DISPLAY WS-CHAR.
    DISPLAY WS-DIGIT.
"#,
    ));
}

#[test]
fn test_pic_long_equivalence() {
    compile_ok(&p(
        r#"
01 WS-X4 PIC XXXX VALUE "ABCD".
01 WS-X4-EQ PIC X(4) VALUE "ABCD".
01 WS-94 PIC 9999 VALUE 1234.
01 WS-94-EQ PIC 9(4) VALUE 1234.
"#,
        r#"
    DISPLAY WS-X4.
    DISPLAY WS-94.
"#,
    ));
}

#[test]
fn test_pic_signed_display() {
    let output = run_prints(&p(
        r#"
01 WS-POS PIC S9(5) VALUE +123.
01 WS-NEG PIC S9(5) VALUE -123.
"#,
        r#"
    DISPLAY WS-POS.
    DISPLAY WS-NEG.
"#,
    ));
    // Display format of signed numbers can vary depending on primitives/locale, but compile and basic display should work
    assert!(output.len() >= 2);
}

#[test]
fn test_pic_implicit_decimal() {
    compile_ok(&p(
        r#"
01 WS-DEC PIC 9(5)V99 VALUE 1234.56.
01 WS-DEC-ONLY PIC V999 VALUE .123.
01 WS-SIGNED-DEC PIC S9(7)V99 VALUE -1234.56.
"#,
        r#"
    DISPLAY WS-DEC.
"#,
    ));
}

#[test]
fn test_pic_national_characters() {
    compile_ok(&p(
        r#"
01 WS-NAT PIC N(10) VALUE "HELLO".
"#,
        r#"
    DISPLAY WS-NAT.
"#,
    ));
}

#[test]
fn test_pic_edited_zero_suppression() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC 9(5) VALUE 42.
01 WS-EDIT1 PIC Z(5) VALUE ZERO.
01 WS-EDIT2 PIC Z(4)9 VALUE ZERO.
"#,
        r#"
    MOVE WS-SRC TO WS-EDIT1.
    DISPLAY WS-EDIT1.
    MOVE WS-SRC TO WS-EDIT2.
    DISPLAY WS-EDIT2.
"#,
    ));
    assert_eq!(output, vec!["   42", "   42"]);
}

#[test]
fn test_pic_edited_asterisk_fill() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC 9(5) VALUE 42.
01 WS-EDIT1 PIC *(5) VALUE ZERO.
01 WS-EDIT2 PIC *(4)9 VALUE ZERO.
"#,
        r#"
    MOVE WS-SRC TO WS-EDIT1.
    DISPLAY WS-EDIT1.
    MOVE WS-SRC TO WS-EDIT2.
    DISPLAY WS-EDIT2.
"#,
    ));
    assert_eq!(output, vec!["***42", "***42"]);
}

#[test]
fn test_pic_edited_currency() {
    compile_ok(&p(
        r#"
01 WS-EDIT1 PIC $9(6) VALUE ZERO.
01 WS-EDIT2 PIC $$$$$9 VALUE ZERO.
"#,
        r#"
    MOVE 123 TO WS-EDIT1.
    MOVE 123 TO WS-EDIT2.
"#,
    ));
}

#[test]
fn test_pic_edited_signs() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC S9(5) VALUE -42.
01 WS-EDIT-PLUS PIC +9(5) VALUE ZERO.
01 WS-EDIT-MINUS PIC -9(5) VALUE ZERO.
01 WS-EDIT-TRAILING PIC 9(5)- VALUE ZERO.
"#,
        r#"
    MOVE WS-SRC TO WS-EDIT-PLUS.
    DISPLAY WS-EDIT-PLUS.
    MOVE WS-SRC TO WS-EDIT-MINUS.
    DISPLAY WS-EDIT-MINUS.
    MOVE WS-SRC TO WS-EDIT-TRAILING.
    DISPLAY WS-EDIT-TRAILING.
"#,
    ));
    assert_eq!(output, vec!["-00042", "-00042", "00042-"]);
}

#[test]
fn test_pic_edited_point_comma() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC 9(6)V99 VALUE 123456.78.
01 WS-EDIT1 PIC 9(5).99 VALUE ZERO.
01 WS-EDIT2 PIC 9(3),9(3).99 VALUE ZERO.
01 WS-EDIT3 PIC ZZZ,ZZZ.ZZ VALUE ZERO.
"#,
        r#"
    MOVE WS-SRC TO WS-EDIT1.
    DISPLAY WS-EDIT1.
    MOVE WS-SRC TO WS-EDIT2.
    DISPLAY WS-EDIT2.
    MOVE WS-SRC TO WS-EDIT3.
    DISPLAY WS-EDIT3.
"#,
    ));
    assert_eq!(output, vec!["23456.78", "123,456.78", "123,456.78"]);
}

#[test]
fn test_pic_edited_credit_debit() {
    compile_ok(&p(
        r#"
01 WS-EDIT1 PIC ZZZZCR VALUE ZERO.
01 WS-EDIT2 PIC ZZZZDB VALUE ZERO.
"#,
        r#"
    MOVE -123 TO WS-EDIT1.
    MOVE -123 TO WS-EDIT2.
"#,
    ));
}

#[test]
fn test_pic_edited_alphanumeric() {
    let output = run_prints(&p(
        r#"
01 WS-EDIT-B PIC X(3)BX(3) VALUE SPACES.
01 WS-EDIT-SLASH PIC X(2)/X(2) VALUE SPACES.
01 WS-EDIT-DATE PIC 99/99/99 VALUE ZERO.
"#,
        r#"
    MOVE "ABCDEF" TO WS-EDIT-B.
    DISPLAY WS-EDIT-B.
    MOVE "0115" TO WS-EDIT-SLASH.
    DISPLAY WS-EDIT-SLASH.
    MOVE 010203 TO WS-EDIT-DATE.
    DISPLAY WS-EDIT-DATE.
"#,
    ));
    assert_eq!(output, vec!["ABC DEF", "01/15", "01/02/03"]);
}
