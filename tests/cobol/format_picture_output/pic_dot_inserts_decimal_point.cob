*> vybe-test: cobol/format_picture_output/pic_dot_inserts_decimal_point
*> origin: languages/cobol/tests/cobol/test_format_picture_output.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(4).99 VALUE 1234.56.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY N.
    MOVE SPACES TO WS-VYBE-L
    STRING N DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "1234.56"
        DISPLAY "FAIL: want [1234.56] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

