*> vybe-test: cobol/reference_modification/test_refmod_target_write
*> origin: languages/cobol/tests/cobol/test_reference_modification.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TXT PIC X(6) VALUE "AABBCC".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    MOVE "XX" TO WS-TXT(3:2).
    DISPLAY WS-TXT.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-TXT DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "AAXXCC"
        DISPLAY "FAIL: want [AAXXCC] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

