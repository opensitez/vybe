*> vybe-test: cobol/add_advanced/test_add_multiple_receiving_fields_runtime
*> origin: languages/cobol/tests/cobol/test_add_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(2) VALUE 10.
01 WS-B PIC 9(2) VALUE 20.
01 WS-C PIC 9(2) VALUE 0.
01 WS-D PIC 9(2) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    ADD WS-A WS-B TO WS-C GIVING WS-D.
    DISPLAY WS-D.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-D DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "30"
        DISPLAY "FAIL: want [30] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

