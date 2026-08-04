*> vybe-test: cobol/string_and_unstring_extended/string_with_multiple_sources_and_delimiters
*> origin: languages/cobol/tests/cobol/test_string_and_unstring_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC X(3) VALUE "A".
01 WS-B PIC X(3) VALUE "B".
01 WS-C PIC X(3) VALUE "C".
01 WS-R PIC X(12) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    STRING WS-A DELIMITED BY SIZE
           WS-B DELIMITED BY SIZE
           WS-C DELIMITED BY SIZE
           INTO WS-R.
    DISPLAY WS-R.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "A  B  C     "
        DISPLAY "FAIL: want [A  B  C     ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

