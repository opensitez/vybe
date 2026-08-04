*> vybe-test: cobol/nested_if_else/if_not_alphabetic_class
*> origin: languages/cobol/tests/cobol/test_nested_if_else.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(5) VALUE "123AB".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF NOT S IS ALPHABETIC
        DISPLAY "MIXED"
    ELSE
        DISPLAY "ALPHA"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "MIXED" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "MIXED"
        DISPLAY "FAIL: want [MIXED] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

