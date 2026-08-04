*> vybe-test: cobol/string_and_unstring_extended/string_with_literal_and_variable_sources
*> origin: languages/cobol/tests/cobol/test_string_and_unstring_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(5) VALUE "COBOL".
01 WS-R PIC X(12) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    STRING "HELLO " DELIMITED BY SIZE
           WS-NAME DELIMITED BY SIZE
           INTO WS-R.
    DISPLAY WS-R.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "HELLO COBOL "
        DISPLAY "FAIL: want [HELLO COBOL ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

