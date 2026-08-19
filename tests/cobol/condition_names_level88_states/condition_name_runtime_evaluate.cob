*> vybe-test: cobol/condition_names_level88_states/condition_name_runtime_evaluate_true_selects_expected_when
*> origin: languages/cobol/tests/cobol/test_condition_names_level88_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC 9 VALUE 1.
   88 ST-A VALUE 1.
   88 ST-B VALUE 2.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE TRUE WHEN ST-B DISPLAY "B" WHEN ST-A DISPLAY "A" WHEN OTHER DISPLAY "X" END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "B" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "B"
        DISPLAY "FAIL: want [B] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

