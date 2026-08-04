*> vybe-test: cobol/reference_modification/test_refmod_in_condition
*> origin: languages/cobol/tests/cobol/test_reference_modification.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TXT PIC X(5) VALUE "ABCDE".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    IF WS-TXT(1:3) = "ABC"
        DISPLAY "MATCH"
    ELSE
        DISPLAY "NO-MATCH"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "MATCH" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "MATCH"
        DISPLAY "FAIL: want [MATCH] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

