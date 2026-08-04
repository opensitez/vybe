*> vybe-test: cobol/paragraph_section_flow/paragraph_modifies_flag_tested_in_main
*> origin: languages/cobol/tests/cobol/test_paragraph_section_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 READY PIC X VALUE "N".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    PERFORM INIT.
    IF READY = "Y"
        DISPLAY "READY"
    ELSE
        DISPLAY "NOT READY"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "READY" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "READY"
        DISPLAY "FAIL: want [READY] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.
INIT.
    MOVE "Y" TO READY.
    STOP RUN.

