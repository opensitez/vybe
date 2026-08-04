*> vybe-test: cobol/special_registers_detail/address_of_elementary_then_display_ws
*> origin: languages/cobol/tests/cobol/test_special_registers_detail.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 DATA-ITEM PIC X(10) VALUE "HELLO".
01 DATA-PTR USAGE POINTER.
PROCEDURE DIVISION.
    SET DATA-PTR TO ADDRESS OF DATA-ITEM.
    DISPLAY DATA-ITEM.
    STOP RUN.

