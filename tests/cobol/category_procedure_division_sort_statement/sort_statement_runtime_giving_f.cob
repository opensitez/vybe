*> vybe-test: cobol/category_procedure_division_sort_statement/sort_statement_runtime_giving_file_compiles
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_sort_statement.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. SRT6.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT S ASSIGN TO "SRT6".
       DATA DIVISION.
       FILE SECTION.
        SD S.
        01 R.
            05 K PIC X(1).
            05 V PIC X(3).
        PROCEDURE DIVISION.
            SORT S
                ON ASCENDING KEY K
                USING "input.dat"
                GIVING "out.dat"
            STOP RUN.

