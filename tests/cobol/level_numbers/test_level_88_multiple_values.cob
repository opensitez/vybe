*> vybe-test: cobol/level_numbers/test_level_88_multiple_values
*> origin: languages/cobol/tests/cobol/test_level_numbers.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-CODE PIC X VALUE "B".
   88 IS-VALID-CODE VALUE "A", "B", "C".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    IF IS-VALID-CODE
        DISPLAY "VALID"
    ELSE
        DISPLAY "INVALID"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "VALID" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "VALID"
        DISPLAY "FAIL: want [VALID] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

