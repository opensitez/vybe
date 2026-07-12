use super::helpers::compile_ok;

// ── ALL figurative constant ───────────────────────────────────

#[test]
fn all_single_char_fill() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-line PIC X(20).
       PROCEDURE DIVISION.
           MOVE ALL "-" TO ws-line
           DISPLAY ws-line
           STOP RUN.
"#,
    );
}

#[test]
fn all_multi_char_pattern() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-banner PIC X(12).
       PROCEDURE DIVISION.
           MOVE ALL "AB" TO ws-banner
           DISPLAY ws-banner
           STOP RUN.
"#,
    );
}

#[test]
fn all_zero_fill() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-field PIC X(10).
       PROCEDURE DIVISION.
           MOVE ALL "0" TO ws-field
           DISPLAY ws-field
           STOP RUN.
"#,
    );
}

#[test]
fn all_star_fill() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-border PIC X(40).
       PROCEDURE DIVISION.
           MOVE ALL "*" TO ws-border
           DISPLAY ws-border
           STOP RUN.
"#,
    );
}

#[test]
fn all_in_string() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-pad PIC X(5).
       01 ws-result PIC X(15).
       PROCEDURE DIVISION.
           MOVE ALL "." TO ws-pad
           STRING "Hi" DELIMITED SIZE
                  ws-pad DELIMITED SIZE
                  INTO ws-result
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

#[test]
fn all_in_compare() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-field PIC X(10) VALUE ALL "*".
       PROCEDURE DIVISION.
           IF ws-field = ALL "*"
               DISPLAY "all stars"
           ELSE
               DISPLAY "not all stars"
           END-IF
           STOP RUN.
"#,
    );
}

#[test]
fn all_numeric_pattern() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-mask PIC X(8).
       PROCEDURE DIVISION.
           MOVE ALL "12" TO ws-mask
           DISPLAY ws-mask
           STOP RUN.
"#,
    );
}

#[test]
fn all_as_initial_value() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-underline PIC X(30) VALUE ALL "-".
       01 ws-dots      PIC X(20) VALUE ALL "...".
       PROCEDURE DIVISION.
           DISPLAY ws-underline
           DISPLAY ws-dots
           STOP RUN.
"#,
    );
}

// ── HIGH-VALUES / HIGH-VALUE ──────────────────────────────────

#[test]
fn high_values_move() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-key PIC X(10).
       PROCEDURE DIVISION.
           MOVE HIGH-VALUES TO ws-key
           DISPLAY "high values set"
           STOP RUN.
"#,
    );
}

#[test]
fn high_value_singular() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-sentinel PIC X.
       PROCEDURE DIVISION.
           MOVE HIGH-VALUE TO ws-sentinel
           IF ws-sentinel = HIGH-VALUE
               DISPLAY "is high value"
           END-IF
           STOP RUN.
"#,
    );
}

#[test]
fn high_values_in_compare() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-key   PIC X(5) VALUE "ZZZZZ".
       01 ws-limit PIC X(5).
       PROCEDURE DIVISION.
           MOVE HIGH-VALUES TO ws-limit
           IF ws-key < ws-limit
               DISPLAY "key is below max"
           END-IF
           STOP RUN.
"#,
    );
}

#[test]
fn high_values_end_of_table_sentinel() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-table.
           05 ws-entry OCCURS 5 TIMES.
               10 ws-code PIC X(5).
       01 ws-idx PIC 9.
       PROCEDURE DIVISION.
           MOVE HIGH-VALUES TO ws-code(1)
           MOVE HIGH-VALUES TO ws-code(2)
           MOVE HIGH-VALUES TO ws-code(3)
           MOVE HIGH-VALUES TO ws-code(4)
           MOVE HIGH-VALUES TO ws-code(5)
           PERFORM VARYING ws-idx FROM 1 BY 1
               UNTIL ws-idx > 5 OR ws-code(ws-idx) = HIGH-VALUES
               DISPLAY ws-code(ws-idx)
           END-PERFORM
           STOP RUN.
"#,
    );
}

#[test]
fn high_values_in_evaluate() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-status PIC X(5).
       PROCEDURE DIVISION.
           MOVE HIGH-VALUES TO ws-status
           EVALUATE ws-status
               WHEN HIGH-VALUES DISPLAY "end of data"
               WHEN LOW-VALUES  DISPLAY "start of data"
               WHEN OTHER       DISPLAY "data: " ws-status
           END-EVALUATE
           STOP RUN.
"#,
    );
}

// ── LOW-VALUES / LOW-VALUE ────────────────────────────────────

#[test]
fn low_values_move() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-init PIC X(10).
       PROCEDURE DIVISION.
           MOVE LOW-VALUES TO ws-init
           DISPLAY "low values set"
           STOP RUN.
"#,
    );
}

#[test]
fn low_value_singular() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-nul PIC X.
       PROCEDURE DIVISION.
           MOVE LOW-VALUE TO ws-nul
           IF ws-nul = LOW-VALUE
               DISPLAY "is low value"
           END-IF
           STOP RUN.
"#,
    );
}

#[test]
fn low_values_as_minimum_key() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-search-key PIC X(10).
       01 ws-result     PIC X VALUE "N".
       PROCEDURE DIVISION.
           MOVE LOW-VALUES TO ws-search-key
           IF ws-search-key < "AAAA"
               MOVE "Y" TO ws-result
           END-IF
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

#[test]
fn low_values_initialize_field() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-buf PIC X(20) VALUE LOW-VALUES.
       PROCEDURE DIVISION.
           DISPLAY "initialized"
           STOP RUN.
"#,
    );
}

// ── QUOTES / QUOTE ────────────────────────────────────────────

#[test]
fn quote_in_string() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-msg PIC X(30).
       PROCEDURE DIVISION.
           MOVE QUOTE TO ws-msg
           DISPLAY ws-msg
           STOP RUN.
"#,
    );
}

#[test]
fn quotes_fill() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-field PIC X(5).
       PROCEDURE DIVISION.
           MOVE QUOTES TO ws-field
           DISPLAY ws-field
           STOP RUN.
"#,
    );
}

// ── ZEROS / ZERO / ZEROES ─────────────────────────────────────

#[test]
fn zeros_to_numeric() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-n PIC 9(5).
       PROCEDURE DIVISION.
           MOVE ZEROS TO ws-n
           DISPLAY ws-n
           STOP RUN.
"#,
    );
}

#[test]
fn zero_singular() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-counter PIC 99 VALUE 5.
       PROCEDURE DIVISION.
           IF ws-counter = ZERO
               DISPLAY "zero"
           ELSE
               DISPLAY "non-zero"
           END-IF
           STOP RUN.
"#,
    );
}

#[test]
fn zeroes_to_alpha() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-field PIC X(6).
       PROCEDURE DIVISION.
           MOVE ZEROES TO ws-field
           DISPLAY ws-field
           STOP RUN.
"#,
    );
}

// ── SPACES / SPACE ────────────────────────────────────────────

#[test]
fn spaces_to_alpha() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-name PIC X(20) VALUE "John Doe".
       PROCEDURE DIVISION.
           MOVE SPACES TO ws-name
           IF ws-name = SPACES
               DISPLAY "cleared"
           END-IF
           STOP RUN.
"#,
    );
}

#[test]
fn space_singular_compare() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-ch PIC X VALUE SPACE.
       PROCEDURE DIVISION.
           IF ws-ch = SPACE
               DISPLAY "blank"
           END-IF
           STOP RUN.
"#,
    );
}

// ── Figurative constants as VALUE clauses ─────────────────────

#[test]
fn value_high_values() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-max-key PIC X(10) VALUE HIGH-VALUES.
       01 ws-min-key PIC X(10) VALUE LOW-VALUES.
       01 ws-blank   PIC X(10) VALUE SPACES.
       01 ws-zero    PIC 9(5)  VALUE ZEROS.
       PROCEDURE DIVISION.
           DISPLAY "initialized"
           STOP RUN.
"#,
    );
}

#[test]
fn figurative_in_condition_all_paths() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-field PIC X(5).
       PROCEDURE DIVISION.
           MOVE "HELLO" TO ws-field
           EVALUATE TRUE
               WHEN ws-field = SPACES     DISPLAY "spaces"
               WHEN ws-field = HIGH-VALUES DISPLAY "high"
               WHEN ws-field = LOW-VALUES  DISPLAY "low"
               WHEN ws-field = ZEROS       DISPLAY "zeros"
               WHEN ws-field = ALL "*"     DISPLAY "stars"
               WHEN OTHER                  DISPLAY ws-field
           END-EVALUATE
           STOP RUN.
"#,
    );
}
