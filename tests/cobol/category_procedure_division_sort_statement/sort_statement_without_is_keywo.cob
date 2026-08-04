*> vybe-test: cobol/category_procedure_division_sort_statement/sort_statement_without_is_keyword_is_accepted
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_sort_statement.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. SRT5.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT S ASSIGN TO "SRT5".
       DATA DIVISION.
       FILE SECTION.
       SD S.
       01 R.
           05 K PIC X(1).
           05 V PIC X(3).
       PROCEDURE DIVISION.
           SORT S
               ON ASCENDING KEY K
               INPUT PROCEDURE SRT-IN
               OUTPUT PROCEDURE SRT-OUT.
           STOP RUN.
       SRT-IN.
       SRT-OUT.

