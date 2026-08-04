*> vybe-test: cobol/condition_names/condition_name_can_gate_nested_if_logic
*> origin: languages/cobol/tests/cobol/test_condition_names.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STATE PIC X VALUE "Y".
   88 IS-OPEN VALUE "Y".
01 WS-COUNT PIC 9 VALUE 1.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    IF IS-OPEN
        IF WS-COUNT > 0
            DISPLAY "OPEN-COUNT"
        END-IF
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "OPEN-COUNT" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OPEN-COUNT"
        DISPLAY "FAIL: want [OPEN-COUNT] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

