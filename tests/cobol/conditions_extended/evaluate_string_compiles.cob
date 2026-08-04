*> vybe-test: cobol/conditions_extended/evaluate_string_compiles
*> origin: languages/cobol/tests/cobol/test_conditions_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC X(1) VALUE "B".
PROCEDURE DIVISION.
    EVALUATE WS-A
        WHEN "A"
            DISPLAY "ALPHA"
        WHEN "B"
            DISPLAY "BETA"
    END-EVALUATE.
    STOP RUN.

