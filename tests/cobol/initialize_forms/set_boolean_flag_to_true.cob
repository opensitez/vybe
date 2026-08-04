*> vybe-test: cobol/initialize_forms/set_boolean_flag_to_true
*> origin: languages/cobol/tests/cobol/test_initialize_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 F PIC X VALUE "N".
    88 F-YES VALUE "Y".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    SET F-YES TO TRUE.
    DISPLAY F.
    MOVE SPACES TO WS-VYBE-L
    STRING F DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "Y"
        DISPLAY "FAIL: want [Y] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

