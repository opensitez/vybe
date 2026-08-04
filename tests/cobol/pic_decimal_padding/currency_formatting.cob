*> vybe-test: cobol/pic_decimal_padding/currency_formatting
*> origin: languages/cobol/tests/cobol/test_pic_decimal_padding.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. CURR.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-PRICE PIC 9(5)V99 VALUE 1234.56.
01 WS-TOTAL PIC 9(8)V99 VALUE 0.
01 WS-QTY   PIC 9(3) VALUE 10.
PROCEDURE DIVISION.
    COMPUTE WS-TOTAL = WS-PRICE * WS-QTY.
    DISPLAY "Unit Price: " WS-PRICE.
    DISPLAY "Quantity:   " WS-QTY.
    DISPLAY "Total:      " WS-TOTAL.
    STOP RUN.

