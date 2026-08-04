*> vybe-test: cobol/category_procedure_division_sort_statement/sort_statement_with_duplicate_control_is_accepted
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_sort_statement.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. SRT4.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT S ASSIGN TO "SRT4".
       DATA DIVISION.
       FILE SECTION.
       SD S.
       01 R.
           05 K PIC X(1).
           05 V PIC X(3).
       PROCEDURE DIVISION.
           SORT S
               ON ASCENDING KEY K
               WITH DUPLICATES IN ORDER
               INPUT PROCEDURE SRT-IN
               OUTPUT PROCEDURE SRT-OUT.
           STOP RUN.
       SRT-IN.
       SRT-OUT.

