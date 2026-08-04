*> vybe-test: cobol/string_delimited_forms/string_size_delimiter_appends_full
*> origin: languages/cobol/tests/cobol/test_string_delimited_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(5) VALUE "HELLO".
01 R PIC X(15) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    STRING A DELIMITED BY SIZE INTO R.
    DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "HELLO          "
        DISPLAY "FAIL: want [HELLO          ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

