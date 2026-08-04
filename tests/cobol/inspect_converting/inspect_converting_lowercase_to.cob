*> vybe-test: cobol/inspect_converting/inspect_converting_lowercase_to_uppercase_partial
*> origin: languages/cobol/tests/cobol/test_inspect_converting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(8) VALUE "abcXYZ".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    INSPECT S CONVERTING "abcdefghij" TO "ABCDEFGHIJ".
    DISPLAY S.
    MOVE SPACES TO WS-VYBE-L
    STRING S DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ABCXYZ  "
        DISPLAY "FAIL: want [ABCXYZ  ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

