*> vybe-test: cobol/category_procedure_division_sort_statement/sort_statement_descending_key_runtime
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_sort_statement.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. SRT2.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT S ASSIGN TO "SRT2".
DATA DIVISION.
FILE SECTION.
SD S.
01 R.
    05 K PIC X(1).
    05 V PIC X(3).
PROCEDURE DIVISION.
    SORT S
        ON DESCENDING KEY K
        INPUT PROCEDURE SRT-IN
        OUTPUT PROCEDURE SRT-OUT.
    STOP RUN.
SRT-IN SECTION.
    MOVE "A" TO R
    RELEASE R
    MOVE "C" TO R
    RELEASE R
    MOVE "B" TO R
    RELEASE R.
SRT-OUT SECTION.
    RETURN S AT END GO TO SRT-OUT-DONE END-RETURN
    DISPLAY K
    GO TO SRT-OUT.
SRT-OUT-DONE.
    DISPLAY "DONE".

