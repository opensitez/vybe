*> vybe-test: cobol/condition_names/condition_name_false_when_unused_setter_is_false
*> origin: languages/cobol/tests/cobol/test_condition_names.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-CODE PIC X VALUE "A".
   88 IS-READY VALUE "A".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    SET IS-READY TO FALSE.
    IF IS-READY
        DISPLAY "READY"
    ELSE
        DISPLAY "NOT-READY"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "READY" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "NOT-READY"
        DISPLAY "FAIL: want [NOT-READY] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

