*> vybe-test: cobol/string_delimited_forms/string_literal_and_variable_mixed
*> origin: languages/cobol/tests/cobol/test_string_delimited_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 NAME PIC X(8) VALUE "ALICE   ".
01 R PIC X(20) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    STRING "HELLO " DELIMITED BY SIZE NAME DELIMITED BY SPACE INTO R.
    DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "HELLO ALICE          "
        DISPLAY "FAIL: want [HELLO ALICE          ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

