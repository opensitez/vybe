*> vybe-test: cobol/pic_decimal_padding/mixed_comp_types
*> origin: languages/cobol/tests/cobol/test_pic_decimal_padding.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. MIXCOMP.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(5) COMP VALUE 100.
01 WS-B PIC 9(5) COMP-3 VALUE 200.
01 WS-C PIC 9(5) USAGE BINARY VALUE 300.
01 WS-D PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    COMPUTE WS-D = WS-A + WS-B + WS-C.
    DISPLAY "Result: " WS-D.
    STOP RUN.

