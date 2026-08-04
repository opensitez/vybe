*> vybe-test: cobol/strings_and_text/reference_modification_extracts_text_slice
*> origin: languages/cobol/tests/cobol/test_strings_and_text.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TXT PIC X(10) VALUE "HELLOTEST".
01 WS-SUB PIC X(4) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE WS-TXT(6:4) TO WS-SUB.
    DISPLAY WS-SUB.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-SUB DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "TEST"
        DISPLAY "FAIL: want [TEST] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

