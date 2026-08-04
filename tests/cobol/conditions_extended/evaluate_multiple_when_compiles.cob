*> vybe-test: cobol/conditions_extended/evaluate_multiple_when_compiles
*> origin: languages/cobol/tests/cobol/test_conditions_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(1) VALUE 3.
PROCEDURE DIVISION.
    EVALUATE WS-A
        WHEN 1
            DISPLAY "ONE"
        WHEN 2
            DISPLAY "TWO"
        WHEN 3
            DISPLAY "THREE"
    END-EVALUATE.
    STOP RUN.

