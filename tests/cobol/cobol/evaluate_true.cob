*> vybe-test: cobol/cobol/evaluate_true
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. EVTRUE.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEMP PIC S9(3) VALUE 25.
PROCEDURE DIVISION.
    EVALUATE TRUE
        WHEN WS-TEMP > 30
            DISPLAY "Hot"
        WHEN WS-TEMP > 20
            DISPLAY "Warm"
        WHEN WS-TEMP > 10
            DISPLAY "Cool"
        WHEN OTHER
            DISPLAY "Cold"
    END-EVALUATE.
    STOP RUN.

