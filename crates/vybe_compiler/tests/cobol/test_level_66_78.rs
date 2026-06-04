use super::helpers::compile_ok;

// ── Level 66 RENAMES ──────────────────────────────────────────

#[test]
fn renames_basic() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 address-record.
           05 street    PIC X(30).
           05 city      PIC X(20).
           05 state     PIC XX.
           05 zip-code  PIC X(10).
       66 city-state RENAMES city THRU state.
       PROCEDURE DIVISION.
           MOVE "Springfield" TO city
           MOVE "IL"          TO state
           DISPLAY city-state
           STOP RUN.
"#,
    );
}

#[test]
fn renames_thru_multiple_fields() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 employee-record.
           05 emp-first  PIC X(15).
           05 emp-middle PIC X(1).
           05 emp-last   PIC X(20).
           05 emp-dept   PIC X(10).
       66 emp-name RENAMES emp-first THRU emp-last.
       PROCEDURE DIVISION.
           MOVE "John"    TO emp-first
           MOVE "A"       TO emp-middle
           MOVE "Smith"   TO emp-last
           DISPLAY emp-name
           STOP RUN.
"#,
    );
}

#[test]
fn renames_single_field() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 payment-rec.
           05 pay-amount    PIC 9(8)V99.
           05 pay-currency  PIC XXX.
       66 pay-curr-alias RENAMES pay-currency.
       PROCEDURE DIVISION.
           MOVE "USD" TO pay-currency
           DISPLAY pay-curr-alias
           STOP RUN.
"#,
    );
}

#[test]
fn renames_numeric_fields() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 date-record.
           05 date-year  PIC 9(4).
           05 date-month PIC 99.
           05 date-day   PIC 99.
       66 date-yymmdd RENAMES date-year THRU date-day.
       PROCEDURE DIVISION.
           MOVE 2024 TO date-year
           MOVE 12   TO date-month
           MOVE 25   TO date-day
           DISPLAY date-yymmdd
           STOP RUN.
"#,
    );
}

#[test]
fn renames_in_move() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 source-rec.
           05 src-a PIC X(5).
           05 src-b PIC X(5).
           05 src-c PIC X(5).
       01 target-rec.
           05 tgt-data PIC X(15).
       66 src-ab RENAMES src-a THRU src-b.
       PROCEDURE DIVISION.
           MOVE "AAAA " TO src-a
           MOVE "BBBBB" TO src-b
           MOVE src-ab TO tgt-data
           DISPLAY tgt-data
           STOP RUN.
"#,
    );
}

#[test]
fn renames_with_redefines_sibling() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 packed-data.
           05 pd-type    PIC X.
           05 pd-code    PIC 9(4).
           05 pd-amount  PIC 9(7)V99.
       66 pd-code-amount RENAMES pd-code THRU pd-amount.
       PROCEDURE DIVISION.
           MOVE "A"  TO pd-type
           MOVE 1234 TO pd-code
           MOVE 9999.99 TO pd-amount
           DISPLAY pd-code-amount
           STOP RUN.
"#,
    );
}

// ── Level 78 constants (COBOL 2002+) ─────────────────────────

#[test]
fn level_78_integer_constant() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       78 MAX-SIZE       VALUE 100.
       78 MIN-SIZE       VALUE 1.
       01 ws-count       PIC 999.
       PROCEDURE DIVISION.
           MOVE MAX-SIZE TO ws-count
           DISPLAY ws-count
           STOP RUN.
"#,
    );
}

#[test]
fn level_78_real_constant() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       78 PI             VALUE 3.14159265.
       78 E              VALUE 2.71828182.
       01 ws-result      PIC 9(3)V9(8).
       PROCEDURE DIVISION.
           MOVE PI TO ws-result
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

#[test]
fn level_78_string_constant() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       78 APP-NAME       VALUE "COBOL Application".
       78 APP-VERSION    VALUE "1.0.0".
       01 ws-title       PIC X(40).
       PROCEDURE DIVISION.
           MOVE APP-NAME TO ws-title
           DISPLAY ws-title
           STOP RUN.
"#,
    );
}

#[test]
fn level_78_used_in_compute() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       78 TAX-RATE       VALUE 0.085.
       78 DISCOUNT-RATE  VALUE 0.10.
       01 ws-price       PIC 9(5)V99 VALUE 100.00.
       01 ws-tax         PIC 9(5)V99.
       01 ws-discount    PIC 9(5)V99.
       PROCEDURE DIVISION.
           COMPUTE ws-tax     = ws-price * TAX-RATE
           COMPUTE ws-discount = ws-price * DISCOUNT-RATE
           DISPLAY ws-tax
           DISPLAY ws-discount
           STOP RUN.
"#,
    );
}

#[test]
fn level_78_used_in_if() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       78 PASS-MARK      VALUE 50.
       78 DISTINCTION    VALUE 85.
       01 ws-score       PIC 999 VALUE 92.
       PROCEDURE DIVISION.
           IF ws-score >= DISTINCTION
               DISPLAY "Distinction"
           ELSE IF ws-score >= PASS-MARK
               DISPLAY "Pass"
           ELSE
               DISPLAY "Fail"
           END-IF
           STOP RUN.
"#,
    );
}

#[test]
fn level_78_multiple_constants() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       78 STATUS-OK      VALUE 0.
       78 STATUS-WARN    VALUE 1.
       78 STATUS-ERROR   VALUE 2.
       78 STATUS-FATAL   VALUE 3.
       01 ws-status      PIC 9.
       PROCEDURE DIVISION.
           MOVE STATUS-OK TO ws-status
           EVALUATE ws-status
               WHEN STATUS-OK    DISPLAY "OK"
               WHEN STATUS-WARN  DISPLAY "Warning"
               WHEN STATUS-ERROR DISPLAY "Error"
               WHEN STATUS-FATAL DISPLAY "Fatal"
           END-EVALUATE
           STOP RUN.
"#,
    );
}

#[test]
fn level_78_in_perform() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       78 LOOP-COUNT     VALUE 5.
       01 ws-idx         PIC 9.
       01 ws-sum         PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           PERFORM VARYING ws-idx FROM 1 BY 1
               UNTIL ws-idx > LOOP-COUNT
               ADD ws-idx TO ws-sum
           END-PERFORM
           DISPLAY ws-sum
           STOP RUN.
"#,
    );
}

#[test]
fn level_78_boolean_constant() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       78 TRUE-VAL       VALUE "Y".
       78 FALSE-VAL      VALUE "N".
       01 ws-active      PIC X VALUE FALSE-VAL.
       PROCEDURE DIVISION.
           MOVE TRUE-VAL TO ws-active
           IF ws-active = TRUE-VAL
               DISPLAY "Active"
           END-IF
           STOP RUN.
"#,
    );
}

#[test]
fn level_78_with_level_88() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       78 MAX-RETRIES    VALUE 3.
       01 ws-retry-count PIC 9 VALUE 0.
           88 max-reached VALUE MAX-RETRIES.
       PROCEDURE DIVISION.
           MOVE MAX-RETRIES TO ws-retry-count
           IF max-reached
               DISPLAY "Max retries reached"
           END-IF
           STOP RUN.
"#,
    );
}

#[test]
fn level_78_mixed_with_66() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       78 RECORD-SIZE    VALUE 80.
       01 full-record.
           05 rec-header PIC X(10).
           05 rec-body   PIC X(60).
           05 rec-footer PIC X(10).
       66 rec-content RENAMES rec-header THRU rec-body.
       01 ws-len PIC 999.
       PROCEDURE DIVISION.
           MOVE RECORD-SIZE TO ws-len
           MOVE "HDR" TO rec-header
           DISPLAY ws-len
           STOP RUN.
"#,
    );
}
