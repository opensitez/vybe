*> vybe-test: cobol/condition_names/condition_name_in_evaluate_selects_matching_branch
*> origin: languages/cobol/tests/cobol/test_condition_names.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STATUS PIC X VALUE "B".
   88 IS-STARTED VALUE "A".
   88 IS-READY VALUE "B".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    EVALUATE TRUE
        WHEN IS-STARTED
            DISPLAY "STARTED"
        WHEN IS-READY
            DISPLAY "READY"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "STARTED" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "STARTED"
        DISPLAY "FAIL: want [STARTED] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

