*> vybe-test: cobol/inspect_converting/inspect_tallying_all_occurrence_count
*> origin: languages/cobol/tests/cobol/test_inspect_converting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(15) VALUE "ABRACADABRA".
01 C PIC 9(3) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    INSPECT S TALLYING C FOR ALL "A".
    DISPLAY C.
    MOVE SPACES TO WS-VYBE-L
    STRING C DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "005"
        DISPLAY "FAIL: want [005] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

