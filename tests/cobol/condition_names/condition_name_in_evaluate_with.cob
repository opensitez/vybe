*> vybe-test: cobol/condition_names/condition_name_in_evaluate_with_other
*> origin: languages/cobol/tests/cobol/test_condition_names.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VAL PIC 99 VALUE 10.
   88 IS-LOW VALUE 1 THRU 5.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    EVALUATE TRUE
        WHEN IS-LOW
            DISPLAY "LOW"
        WHEN OTHER
            DISPLAY "HIGH"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "LOW" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "LOW"
        DISPLAY "FAIL: want [LOW] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

