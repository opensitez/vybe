*> vybe-test: cobol/pic_decimal_padding/padded_record_output
*> origin: languages/cobol/tests/cobol/test_pic_decimal_padding.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. PADREC.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-RECORD.
   05 WS-ID     PIC 9(5)  VALUE 0.
   05 WS-NAME   PIC X(30) VALUE SPACES.
   05 WS-AMOUNT PIC 9(8)V99 VALUE 0.
PROCEDURE DIVISION.
    MOVE 12345 TO WS-ID.
    MOVE "John Smith" TO WS-NAME.
    MOVE 5000.50 TO WS-AMOUNT.
    DISPLAY WS-ID " " WS-NAME " " WS-AMOUNT.
    STOP RUN.

