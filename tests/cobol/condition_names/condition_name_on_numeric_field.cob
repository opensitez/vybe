*> vybe-test: cobol/condition_names/condition_name_on_numeric_field_is_true_for_range
*> origin: languages/cobol/tests/cobol/test_condition_names.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SCORE PIC 99 VALUE 60.
   88 IS-PASSING VALUE 50 THRU 100.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    IF IS-PASSING
        DISPLAY "PASS"
    ELSE
        DISPLAY "FAIL"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "PASS" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "PASS"
        DISPLAY "FAIL: want [PASS] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

