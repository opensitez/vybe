*> vybe-test: cobol/new_features/exit_perform_program
*> origin: languages/cobol/tests/cobol/test_new_features.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. EXITPERF.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-I PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 100
        IF WS-I = 50
            EXIT PERFORM
        END-IF
        DISPLAY WS-I
    END-PERFORM.
    DISPLAY "Finished at " WS-I.
    STOP RUN.

