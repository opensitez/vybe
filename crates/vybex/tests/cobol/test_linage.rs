use super::helpers::compile_ok;

// ── LINAGE basic ──────────────────────────────────────────────

#[test] fn linage_basic() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT print-file ASSIGN TO "report.txt"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD print-file
           LINAGE IS 20 LINES.
       01 print-rec PIC X(80).
       PROCEDURE DIVISION.
           OPEN OUTPUT print-file
           MOVE "Hello, Report!" TO print-rec
           WRITE print-rec
           CLOSE print-file
           STOP RUN.
"#);
}

#[test] fn linage_with_footing() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT rpt ASSIGN TO "rpt.txt"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD rpt
           LINAGE IS 56 LINES
           WITH FOOTING AT 54.
       01 rpt-line PIC X(132).
       PROCEDURE DIVISION.
           OPEN OUTPUT rpt
           MOVE "Report line 1" TO rpt-line
           WRITE rpt-line
           MOVE "Report line 2" TO rpt-line
           WRITE rpt-line
           CLOSE rpt
           STOP RUN.
"#);
}

#[test] fn linage_with_top_bottom() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT page-file ASSIGN TO "page.txt"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD page-file
           LINAGE IS 40 LINES
           WITH FOOTING AT 38
           LINES AT TOP 3
           LINES AT BOTTOM 3.
       01 page-rec PIC X(80).
       PROCEDURE DIVISION.
           OPEN OUTPUT page-file
           MOVE "First line" TO page-rec
           WRITE page-rec
           CLOSE page-file
           STOP RUN.
"#);
}

#[test] fn linage_lines_at_top_only() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT out-file ASSIGN TO "out.txt"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD out-file
           LINAGE IS 50 LINES
           LINES AT TOP 5.
       01 out-line PIC X(80).
       PROCEDURE DIVISION.
           OPEN OUTPUT out-file
           MOVE "Body line" TO out-line
           WRITE out-line
           CLOSE out-file
           STOP RUN.
"#);
}

#[test] fn linage_lines_at_bottom_only() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT rpt-file ASSIGN TO "rpt.txt"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD rpt-file
           LINAGE IS 60 LINES
           LINES AT BOTTOM 4.
       01 rpt-line PIC X(132).
       PROCEDURE DIVISION.
           OPEN OUTPUT rpt-file
           MOVE "Detail line" TO rpt-line
           WRITE rpt-line
           CLOSE rpt-file
           STOP RUN.
"#);
}

#[test] fn linage_full_specification() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT full-rpt ASSIGN TO "full.txt"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD full-rpt
           LINAGE IS 60 LINES
           WITH FOOTING AT 58
           LINES AT TOP 3
           LINES AT BOTTOM 3.
       01 full-line PIC X(132).
       PROCEDURE DIVISION.
           OPEN OUTPUT full-rpt
           MOVE "Report header line" TO full-line
           WRITE full-line AFTER ADVANCING PAGE
           MOVE "Detail line 1" TO full-line
           WRITE full-line AFTER ADVANCING 1 LINE
           MOVE "Detail line 2" TO full-line
           WRITE full-line AFTER ADVANCING 1 LINE
           CLOSE full-rpt
           STOP RUN.
"#);
}

#[test] fn linage_with_variable() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT var-rpt ASSIGN TO "var.txt"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD var-rpt
           LINAGE IS ws-page-lines LINES
           WITH FOOTING AT ws-footing-line
           LINES AT TOP ws-top-margin
           LINES AT BOTTOM ws-bot-margin.
       01 var-line PIC X(80).
       WORKING-STORAGE SECTION.
       01 ws-page-lines    PIC 99 VALUE 55.
       01 ws-footing-line  PIC 99 VALUE 53.
       01 ws-top-margin    PIC 9  VALUE 2.
       01 ws-bot-margin    PIC 9  VALUE 3.
       PROCEDURE DIVISION.
           OPEN OUTPUT var-rpt
           MOVE "Variable linage test" TO var-line
           WRITE var-line
           CLOSE var-rpt
           STOP RUN.
"#);
}

#[test] fn linage_write_advancing_page() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT pg-file ASSIGN TO "pg.txt"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD pg-file
           LINAGE IS 30 LINES
           WITH FOOTING AT 28.
       01 pg-line PIC X(80).
       PROCEDURE DIVISION.
           OPEN OUTPUT pg-file
           MOVE "Page 1 Header" TO pg-line
           WRITE pg-line AFTER ADVANCING PAGE
           MOVE "Page 1 body"   TO pg-line
           WRITE pg-line AFTER ADVANCING 1 LINE
           MOVE "Page 2 Header" TO pg-line
           WRITE pg-line AFTER ADVANCING PAGE
           CLOSE pg-file
           STOP RUN.
"#);
}

#[test] fn linage_counter_access() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT cnt-file ASSIGN TO "cnt.txt"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD cnt-file
           LINAGE IS 25 LINES
           WITH FOOTING AT 23.
       01 cnt-line PIC X(80).
       WORKING-STORAGE SECTION.
       01 ws-line-no PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           OPEN OUTPUT cnt-file
           PERFORM 10 TIMES
               MOVE LINAGE-COUNTER TO ws-line-no
               STRING "Line " DELIMITED SIZE
                      ws-line-no DELIMITED SIZE
                      INTO cnt-line
               WRITE cnt-line
           END-PERFORM
           CLOSE cnt-file
           STOP RUN.
"#);
}

#[test] fn linage_at_end_of_page() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT eop-file ASSIGN TO "eop.txt"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD eop-file
           LINAGE IS 10 LINES
           WITH FOOTING AT 9.
       01 eop-line PIC X(80).
       WORKING-STORAGE SECTION.
       01 ws-idx PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           OPEN OUTPUT eop-file
           PERFORM VARYING ws-idx FROM 1 BY 1 UNTIL ws-idx > 25
               MOVE ws-idx TO eop-line
               WRITE eop-line
                   AT END-OF-PAGE
                       MOVE "--- page break ---" TO eop-line
                       WRITE eop-line AFTER ADVANCING PAGE
               END-WRITE
           END-PERFORM
           CLOSE eop-file
           STOP RUN.
"#);
}

#[test] fn linage_not_at_end_of_page() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT nep-file ASSIGN TO "nep.txt"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD nep-file
           LINAGE IS 20 LINES
           WITH FOOTING AT 18.
       01 nep-line PIC X(80).
       WORKING-STORAGE SECTION.
       01 ws-line-count PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           OPEN OUTPUT nep-file
           MOVE "Test line" TO nep-line
           WRITE nep-line
               AT END-OF-PAGE
                   ADD 1 TO ws-line-count
               NOT AT END-OF-PAGE
                   DISPLAY "still on page"
           END-WRITE
           CLOSE nep-file
           STOP RUN.
"#);
}
