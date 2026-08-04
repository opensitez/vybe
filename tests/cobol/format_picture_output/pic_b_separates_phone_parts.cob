*> vybe-test: cobol/format_picture_output/pic_b_separates_phone_parts
*> origin: languages/cobol/tests/cobol/test_format_picture_output.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 PHONE PIC 9(3)BB9(3)BB9(4) VALUE 5551234567.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY PHONE.
    MOVE SPACES TO WS-VYBE-L
    STRING PHONE DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "555  123  4567"
        DISPLAY "FAIL: want [555  123  4567] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

