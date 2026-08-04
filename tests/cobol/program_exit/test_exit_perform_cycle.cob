*> vybe-test: cobol/program_exit/test_exit_perform_cycle
*> origin: languages/cobol/tests/cobol/test_program_exit.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-I PIC 9.
PROCEDURE DIVISION.

    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 5
        IF WS-I = 3
            EXIT PERFORM CYCLE
        END-IF
        DISPLAY WS-I
    END-PERFORM.
    STOP RUN.

