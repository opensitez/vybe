*> vybe-test: cobol/string_delimited_forms/string_with_five_literals
*> origin: languages/cobol/tests/cobol/test_string_delimited_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC X(20) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    STRING "A" DELIMITED BY SIZE "B" DELIMITED BY SIZE "C" DELIMITED BY SIZE "D" DELIMITED BY SIZE "E" DELIMITED BY SIZE INTO R.
    DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ABCDE               "
        DISPLAY "FAIL: want [ABCDE               ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

