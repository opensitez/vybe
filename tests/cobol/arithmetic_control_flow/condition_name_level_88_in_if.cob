*> vybe-test: cobol/arithmetic_control_flow/condition_name_level_88_in_if
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-S PIC X VALUE "A".
   88 IS-A VALUE "A".
   88 IS-B VALUE "B".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF IS-A
        DISPLAY "A-STATE"
    ELSE
        DISPLAY "OTHER"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "A-STATE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "A-STATE"
        DISPLAY "FAIL: want [A-STATE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

