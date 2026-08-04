*> vybe-test: cobol/initialize_forms/set_condition_false_and_verify
*> origin: languages/cobol/tests/cobol/test_initialize_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 F PIC X VALUE "Y".
    88 F-ON VALUE "Y".
    88 F-OFF VALUE "N".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    SET F-ON TO FALSE.
    IF F-OFF
        DISPLAY "OFF"
    ELSE
        DISPLAY "ON"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "OFF" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OFF"
        DISPLAY "FAIL: want [OFF] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

