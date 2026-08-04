*> vybe-test: cobol/conditions/test_condition_class_alphabetic
*> origin: languages/cobol/tests/cobol/test_conditions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TXT PIC X(3) VALUE "ABC".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    IF WS-TXT IS ALPHABETIC
        DISPLAY "ALPHA"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "ALPHA" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ALPHA"
        DISPLAY "FAIL: want [ALPHA] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

