*> vybe-test: cobol/condition_compound/condition_less_than_string
*> origin: languages/cobol/tests/cobol/test_condition_compound.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(3) VALUE "ABC".
01 B PIC X(3) VALUE "DEF".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF A < B
        DISPLAY "A BEFORE B"
    ELSE
        DISPLAY "B BEFORE A"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "A BEFORE B" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "A BEFORE B"
        DISPLAY "FAIL: want [A BEFORE B] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

