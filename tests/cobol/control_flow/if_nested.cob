*> vybe-test: cobol/control_flow/if_nested
*> origin: languages/cobol/tests/cobol/test_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9(3) VALUE 5.
PROCEDURE DIVISION.
    IF X > 10
        DISPLAY "A"
    ELSE
        IF X > 5
            DISPLAY "B"
        ELSE
            DISPLAY "C"
        END-IF
    END-IF.
    STOP RUN.

