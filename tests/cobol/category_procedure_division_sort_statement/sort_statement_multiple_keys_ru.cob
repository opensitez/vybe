*> vybe-test: cobol/category_procedure_division_sort_statement/sort_statement_multiple_keys_runtime
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_sort_statement.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. SRT3.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT S ASSIGN TO "SRT3".
DATA DIVISION.
FILE SECTION.
SD S.
01 R.
    05 K1 PIC X(1).
    05 K2 PIC 9.
    05 V  PIC X(2).
PROCEDURE DIVISION.
    SORT S
        ON ASCENDING KEY K1
        ON DESCENDING KEY K2
        INPUT PROCEDURE SRT-IN
        OUTPUT PROCEDURE SRT-OUT.
    STOP RUN.
SRT-IN SECTION.
    MOVE "A2" TO R
    RELEASE R
    MOVE "A9" TO R
    RELEASE R
    MOVE "B1" TO R
    RELEASE R.
SRT-OUT SECTION.
    RETURN S
        AT END DISPLAY "DONE"
        NOT AT END
            DISPLAY K1
            DISPLAY K2
        END-RETURN.

