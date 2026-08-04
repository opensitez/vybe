*> vybe-test: cobol/format_picture_output/pic_b_inserts_blank
*> origin: languages/cobol/tests/cobol/test_format_picture_output.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(3)B9(3) VALUE 123456.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY N.
    MOVE SPACES TO WS-VYBE-L
    STRING N DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "123 456"
        DISPLAY "FAIL: want [123 456] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

