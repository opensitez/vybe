use super::helpers::compile_ok;

// ── ROUNDED (default — nearest away from zero) ─────────────────

#[test]
fn rounded_basic_add() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9V9 VALUE 0.
       PROCEDURE DIVISION.
           COMPUTE ws-result ROUNDED = 1.35
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

#[test]
fn rounded_basic_subtract() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-a      PIC 9V99 VALUE 5.67.
       01 ws-b      PIC 9V99 VALUE 2.34.
       01 ws-result PIC 9V9  VALUE 0.
       PROCEDURE DIVISION.
           SUBTRACT ws-b FROM ws-a GIVING ws-result ROUNDED
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

#[test]
fn rounded_basic_multiply() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           MULTIPLY 3 BY 3.7 GIVING ws-result ROUNDED
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

#[test]
fn rounded_basic_divide() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9V99 VALUE 0.
       PROCEDURE DIVISION.
           DIVIDE 3 INTO 10 GIVING ws-result ROUNDED
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

// ── ROUNDED MODE TRUNCATION ───────────────────────────────────

#[test]
fn rounded_mode_truncation_positive() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9V9 VALUE 0.
       PROCEDURE DIVISION.
           COMPUTE ws-result ROUNDED MODE TRUNCATION = 2.79
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

#[test]
fn rounded_mode_truncation_negative() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC S9V9 VALUE 0.
       PROCEDURE DIVISION.
           COMPUTE ws-result ROUNDED MODE TRUNCATION = -2.79
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

#[test]
fn rounded_mode_truncation_divide() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-q PIC 9V99 VALUE 0.
       PROCEDURE DIVISION.
           DIVIDE 6 INTO 10 GIVING ws-q ROUNDED MODE TRUNCATION
           DISPLAY ws-q
           STOP RUN.
"#,
    );
}

// ── ROUNDED MODE NEAREST-EVEN (banker's rounding) ─────────────

#[test]
fn rounded_mode_nearest_even_half_up() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           COMPUTE ws-result ROUNDED MODE NEAREST-EVEN = 2.5
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

#[test]
fn rounded_mode_nearest_even_half_down() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           COMPUTE ws-result ROUNDED MODE NEAREST-EVEN = 3.5
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

#[test]
fn rounded_mode_nearest_even_divide() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9V99 VALUE 0.
       PROCEDURE DIVISION.
           DIVIDE 3 INTO 1 GIVING ws-result ROUNDED MODE NEAREST-EVEN
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

// ── ROUNDED MODE TOWARD-GREATER (ceiling) ─────────────────────

#[test]
fn rounded_mode_toward_greater_positive() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           COMPUTE ws-result ROUNDED MODE TOWARD-GREATER = 2.1
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

#[test]
fn rounded_mode_toward_greater_negative() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC S9 VALUE 0.
       PROCEDURE DIVISION.
           COMPUTE ws-result ROUNDED MODE TOWARD-GREATER = -2.9
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

#[test]
fn rounded_mode_toward_greater_exact() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           COMPUTE ws-result ROUNDED MODE TOWARD-GREATER = 3.0
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

// ── ROUNDED MODE TOWARD-LESSER (floor) ────────────────────────

#[test]
fn rounded_mode_toward_lesser_positive() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           COMPUTE ws-result ROUNDED MODE TOWARD-LESSER = 2.9
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

#[test]
fn rounded_mode_toward_lesser_negative() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC S9 VALUE 0.
       PROCEDURE DIVISION.
           COMPUTE ws-result ROUNDED MODE TOWARD-LESSER = -2.1
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

// ── ROUNDED MODE NEAREST-TOWARD-ZERO ─────────────────────────

#[test]
fn rounded_mode_nearest_toward_zero_pos() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           COMPUTE ws-result ROUNDED MODE NEAREST-TOWARD-ZERO = 2.5
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

#[test]
fn rounded_mode_nearest_toward_zero_neg() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC S9 VALUE 0.
       PROCEDURE DIVISION.
           COMPUTE ws-result ROUNDED MODE NEAREST-TOWARD-ZERO = -2.5
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

// ── ROUNDED MODE AWAY-FROM-ZERO ───────────────────────────────

#[test]
fn rounded_mode_away_from_zero_pos() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           COMPUTE ws-result ROUNDED MODE AWAY-FROM-ZERO = 2.5
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

#[test]
fn rounded_mode_away_from_zero_neg() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC S9 VALUE 0.
       PROCEDURE DIVISION.
           COMPUTE ws-result ROUNDED MODE AWAY-FROM-ZERO = -2.5
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

// ── ROUNDED MODE PROHIBITED ───────────────────────────────────

#[test]
fn rounded_mode_prohibited_exact() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9V9 VALUE 0.
       01 ws-err    PIC X   VALUE "N".
       PROCEDURE DIVISION.
           COMPUTE ws-result ROUNDED MODE PROHIBITED = 2.5
               ON SIZE ERROR MOVE "Y" TO ws-err
           END-COMPUTE
           DISPLAY ws-err
           STOP RUN.
"#,
    );
}

// ── Multiple ROUNDED in one statement ────────────────────────

#[test]
fn rounded_multiple_giving() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-a   PIC 9V9 VALUE 3.7.
       01 ws-b   PIC 9V9 VALUE 3.7.
       01 ws-c   PIC 9V9 VALUE 3.7.
       PROCEDURE DIVISION.
           ADD 1.25 TO ws-a ROUNDED
           ADD 1.25 TO ws-b ROUNDED MODE TRUNCATION
           ADD 1.25 TO ws-c ROUNDED MODE NEAREST-EVEN
           DISPLAY ws-a
           DISPLAY ws-b
           DISPLAY ws-c
           STOP RUN.
"#,
    );
}

#[test]
fn rounded_in_financial_calc() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-principal  PIC 9(7)V99 VALUE 1000.00.
       01 ws-rate       PIC V9(4)   VALUE .0875.
       01 ws-interest   PIC 9(5)V99 VALUE 0.
       01 ws-total      PIC 9(7)V99 VALUE 0.
       PROCEDURE DIVISION.
           COMPUTE ws-interest ROUNDED MODE NEAREST-EVEN
               = ws-principal * ws-rate
           COMPUTE ws-total = ws-principal + ws-interest
           DISPLAY ws-interest
           DISPLAY ws-total
           STOP RUN.
"#,
    );
}
