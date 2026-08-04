*> vybe-test: cobol/condition_names/condition_name_supports_nested_boolean_checks
*> origin: languages/cobol/tests/cobol/test_condition_names.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STATE PIC X VALUE "Y".
   88 IS-ACTIVE VALUE "Y".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    IF IS-ACTIVE
        DISPLAY "ACTIVE"
    ELSE
        DISPLAY "INACTIVE"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "ACTIVE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ACTIVE"
        DISPLAY "FAIL: want [ACTIVE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

