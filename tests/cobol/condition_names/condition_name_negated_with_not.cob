*> vybe-test: cobol/condition_names/condition_name_negated_with_not
*> origin: languages/cobol/tests/cobol/test_condition_names.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STATE PIC X VALUE "X".
   88 IS-ACTIVE VALUE "Y", "Z".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    IF NOT IS-ACTIVE
        DISPLAY "INACTIVE"
    ELSE
        DISPLAY "ACTIVE"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "INACTIVE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "INACTIVE"
        DISPLAY "FAIL: want [INACTIVE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

