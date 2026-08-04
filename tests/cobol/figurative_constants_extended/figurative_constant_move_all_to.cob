*> vybe-test: cobol/figurative_constants_extended/figurative_constant_move_all_to_numeric_field
*> origin: languages/cobol/tests/cobol/test_figurative_constants_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NUM PIC 9(4) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    MOVE ALL "9" TO WS-NUM.
    DISPLAY WS-NUM.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-NUM DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "9999"
        DISPLAY "FAIL: want [9999] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

