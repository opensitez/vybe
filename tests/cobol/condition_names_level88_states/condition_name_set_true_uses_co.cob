*> vybe-test: cobol/condition_names_level88_states/condition_name_set_true_uses_condition_value
*> origin: languages/cobol/tests/cobol/test_condition_names_level88_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 ST PIC 9 VALUE 0.
   88 ACTIVE VALUE 7.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    SET ACTIVE TO TRUE.
    DISPLAY ST.
    MOVE SPACES TO WS-VYBE-L
    STRING ST DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "7"
        DISPLAY "FAIL: want [7] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

