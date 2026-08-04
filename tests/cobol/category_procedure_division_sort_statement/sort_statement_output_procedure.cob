*> vybe-test: cobol/category_procedure_division_sort_statement/sort_statement_output_procedure_without_is_keyword
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_sort_statement.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. SRT7.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT S ASSIGN TO "SRT7".
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
    MOVE "X" TO R
    RELEASE R.
SRT-OUT.
    RETURN S
        AT END DISPLAY "DONE"
        NOT AT END DISPLAY K
        END-RETURN.

