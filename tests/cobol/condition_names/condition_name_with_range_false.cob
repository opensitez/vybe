*> vybe-test: cobol/condition_names/condition_name_with_range_false_outside_bounds
*> origin: languages/cobol/tests/cobol/test_condition_names.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-AGE PIC 99 VALUE 40.
   88 IS-YOUTH VALUE 15 THRU 30.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    IF IS-YOUTH
        DISPLAY "YOUTH"
    ELSE
        DISPLAY "OTHER"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "YOUTH" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OTHER"
        DISPLAY "FAIL: want [OTHER] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

