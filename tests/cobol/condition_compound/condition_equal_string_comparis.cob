*> vybe-test: cobol/condition_compound/condition_equal_string_comparison
*> origin: languages/cobol/tests/cobol/test_condition_compound.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(5) VALUE "HELLO".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF S = "HELLO"
        DISPLAY "MATCH"
    ELSE
        DISPLAY "NO MATCH"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "MATCH" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "MATCH"
        DISPLAY "FAIL: want [MATCH] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

