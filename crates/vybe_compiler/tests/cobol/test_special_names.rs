use super::helpers::compile_ok;

// ── CURRENCY SIGN ─────────────────────────────────────────────

#[test]
fn currency_sign_default() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           CURRENCY SIGN IS "$".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-amount PIC $9,999.99 VALUE 1234.56.
       PROCEDURE DIVISION.
           DISPLAY ws-amount
           STOP RUN.
"#,
    );
}

#[test]
fn currency_sign_euro() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           CURRENCY SIGN IS "E".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-price PIC E9(6)V99 VALUE 1000.00.
       PROCEDURE DIVISION.
           DISPLAY ws-price
           STOP RUN.
"#,
    );
}

#[test]
fn currency_sign_pound() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           CURRENCY SIGN IS "L".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-amount PIC L9(5)V99 VALUE 500.00.
       PROCEDURE DIVISION.
           DISPLAY ws-amount
           STOP RUN.
"#,
    );
}

// ── DECIMAL-POINT IS COMMA ────────────────────────────────────

#[test]
fn decimal_point_is_comma() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           DECIMAL-POINT IS COMMA.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-val PIC 9.999 VALUE 3,14159.
       PROCEDURE DIVISION.
           DISPLAY ws-val
           STOP RUN.
"#,
    );
}

#[test]
fn decimal_point_comma_arithmetic() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           DECIMAL-POINT IS COMMA.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-a PIC 9V99 VALUE 1,50.
       01 ws-b PIC 9V99 VALUE 2,75.
       01 ws-c PIC 9V99 VALUE 0.
       PROCEDURE DIVISION.
           ADD ws-a ws-b GIVING ws-c
           DISPLAY ws-c
           STOP RUN.
"#,
    );
}

#[test]
fn decimal_point_comma_with_currency() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           DECIMAL-POINT IS COMMA
           CURRENCY SIGN IS "E".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-price PIC E9.999,99 VALUE 1234,56.
       PROCEDURE DIVISION.
           DISPLAY ws-price
           STOP RUN.
"#,
    );
}

// ── CLASS definition ──────────────────────────────────────────

#[test]
fn class_digit_chars() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           CLASS DIGIT-CHARS IS "0" THRU "9".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-char PIC X VALUE "5".
       PROCEDURE DIVISION.
           IF ws-char IS DIGIT-CHARS
               DISPLAY "is digit"
           ELSE
               DISPLAY "not digit"
           END-IF
           STOP RUN.
"#,
    );
}

#[test]
fn class_alpha_chars() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           CLASS UPPER-ALPHA IS "A" THRU "Z"
           CLASS LOWER-ALPHA IS "a" THRU "z".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-ch PIC X VALUE "M".
       PROCEDURE DIVISION.
           EVALUATE TRUE
               WHEN ws-ch IS UPPER-ALPHA DISPLAY "upper"
               WHEN ws-ch IS LOWER-ALPHA DISPLAY "lower"
               WHEN OTHER                DISPLAY "other"
           END-EVALUATE
           STOP RUN.
"#,
    );
}

#[test]
fn class_special_chars() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           CLASS HEX-CHARS IS "0" THRU "9" "A" THRU "F" "a" THRU "f".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-hex PIC X VALUE "F".
       PROCEDURE DIVISION.
           IF ws-hex IS HEX-CHARS
               DISPLAY "valid hex"
           END-IF
           STOP RUN.
"#,
    );
}

#[test]
fn class_with_multiple_ranges() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           CLASS ALNUM IS "0" THRU "9" "A" THRU "Z" "a" THRU "z".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-test PIC X VALUE "X".
       PROCEDURE DIVISION.
           IF ws-test IS ALNUM
               DISPLAY "alphanumeric"
           END-IF
           STOP RUN.
"#,
    );
}

// ── ALPHABET definition ───────────────────────────────────────

#[test]
fn alphabet_ebcdic() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           ALPHABET EBCDIC-ALPHA IS EBCDIC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-char PIC X VALUE "A".
       PROCEDURE DIVISION.
           DISPLAY ws-char
           STOP RUN.
"#,
    );
}

#[test]
fn alphabet_ascii() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           ALPHABET ASCII-ALPHA IS ASCII.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-val PIC X VALUE "Z".
       PROCEDURE DIVISION.
           DISPLAY ws-val
           STOP RUN.
"#,
    );
}

#[test]
fn alphabet_standard_1() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           ALPHABET STD-ALPHA IS STANDARD-1.
       PROCEDURE DIVISION.
           DISPLAY "ok"
           STOP RUN.
"#,
    );
}

#[test]
fn alphabet_user_defined() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           ALPHABET MY-ALPHA IS "A" "B" "C" "D" THRU "Z".
       PROCEDURE DIVISION.
           DISPLAY "ok"
           STOP RUN.
"#,
    );
}

// ── SYMBOLIC CHARACTERS ───────────────────────────────────────

#[test]
fn symbolic_characters_basic() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           SYMBOLIC CHARACTERS TAB IS 9
                               ESC IS 27.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-tab PIC X VALUE TAB.
       PROCEDURE DIVISION.
           DISPLAY "tab char defined"
           STOP RUN.
"#,
    );
}

#[test]
fn symbolic_characters_multiple() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           SYMBOLIC CHARACTERS NULL-CHAR IS 1
                               BELL      IS 7
                               CR-CHAR   IS 13
                               LF-CHAR   IS 10.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-cr PIC X VALUE CR-CHAR.
       PROCEDURE DIVISION.
           DISPLAY "control chars defined"
           STOP RUN.
"#,
    );
}

// ── CONSOLE IS CRT (interactive I/O) ─────────────────────────

#[test]
fn console_is_crt() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           CONSOLE IS CRT.
       PROCEDURE DIVISION.
           DISPLAY "console output"
           STOP RUN.
"#,
    );
}

// ── Combined SPECIAL-NAMES ────────────────────────────────────

#[test]
fn special_names_combined() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           CURRENCY SIGN IS "$"
           DECIMAL-POINT IS COMMA
           CLASS DIGIT IS "0" THRU "9"
           ALPHABET MY-COLL IS STANDARD-1.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-price PIC $9.999,99 VALUE 1234,56.
       01 ws-ch    PIC X VALUE "7".
       PROCEDURE DIVISION.
           IF ws-ch IS DIGIT
               DISPLAY "digit"
           END-IF
           DISPLAY ws-price
           STOP RUN.
"#,
    );
}
