*> vybe-test: cobol/special_registers_detail/address_of_group_item_compiles
*> origin: languages/cobol/tests/cobol/test_special_registers_detail.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 GRP.
   05 F1 PIC X(5) VALUE "HELLO".
01 PTR USAGE POINTER.
PROCEDURE DIVISION.
    SET PTR TO ADDRESS OF GRP.
    STOP RUN.

