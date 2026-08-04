*> vybe-test: cobol/collating_sequence/collating_sequence_user_defined_runtime
*> origin: languages/cobol/tests/cobol/test_collating_sequence.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. COLL11.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    ALPHABET MY-COLLATE IS "B" "A" "C".
    COLLATING SEQUENCE IS MY-COLLATE.
PROCEDURE DIVISION.
    IF "B" < "A"
        DISPLAY "B_BEFORE_A"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "B_BEFORE_A" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "B_BEFORE_A"
        DISPLAY "FAIL: want [B_BEFORE_A] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

