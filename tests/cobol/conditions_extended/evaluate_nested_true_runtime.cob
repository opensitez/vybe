*> vybe-test: cobol/conditions_extended/evaluate_nested_true_runtime
*> origin: languages/cobol/tests/cobol/test_conditions_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VAL PIC X(1) VALUE "C".
01 WS-COUNT PIC 9 VALUE 0.
   88 MATCH VALUE "A".
   88 NO-MATCH VALUE "B".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF WS-COUNT = 0
        EVALUATE TRUE
            WHEN MATCH
                DISPLAY "MATCH"
            WHEN NO-MATCH
                DISPLAY "NO"
            WHEN OTHER
                DISPLAY "OTHER"
        END-EVALUATE
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "MATCH" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OTHER"
        DISPLAY "FAIL: want [OTHER] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

